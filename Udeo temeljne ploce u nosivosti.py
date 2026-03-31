import sys
import numpy as np

sys.path.append(r"C:\ProgramData\Seequent\PLAXIS Python Distribution V3\python\Lib\site-packages")

from plxscripting.easy import *
from openpyxl import Workbook


# ============================================================
# SETTINGS
# ============================================================
HOST = "localhost"
PORT = 10001
PASSWORD = "12345"

OUT_FILE = r"plate_interface_effective_sigma_n_resultant.xlsx"

# Optional filter:
# Leave empty [] to process all interfaces
# Example: TARGET_INTERFACE_NAMES = ["Plate_1", "Wall A"]
TARGET_INTERFACE_NAMES = []


# ============================================================
# CONNECT TO PLAXIS OUTPUT
# ============================================================
s_o, g_o = new_server(HOST, PORT, password=PASSWORD)


# ============================================================
# HELPERS
# ============================================================
def safe_name(obj, fallback):
    try:
        return obj.Name.value
    except:
        return fallback


def safe_float(val):
    try:
        return float(val)
    except:
        return None


def get_last_phase(g_o):
    phase_names = [a for a in dir(g_o) if a.startswith("Phase_")]
    if phase_names:
        phase_names.sort(key=lambda x: int(x.split("_")[1]))
        return getattr(g_o, phase_names[-1])

    if hasattr(g_o, "Phases"):
        return g_o.Phases[-1]

    raise Exception("Cannot find phases in Output.")


def get_all_interfaces(g_o):
    """
    Tries common collection names used by PLAXIS Output.
    """
    candidates = [
        "Interfaces",
        "PositiveInterfaces",
        "NegativeInterfaces",
        "PosInterfaces",
        "NegInterfaces",
    ]

    found = []
    seen = set()

    for c in candidates:
        if hasattr(g_o, c):
            try:
                objs = list(getattr(g_o, c))
                for obj in objs:
                    key = str(obj)
                    if key not in seen:
                        seen.add(key)
                        found.append(obj)
            except:
                pass

    if not found:
        raise Exception(
            "No interface collections found. Check dir(g_o) to see how interfaces are exposed in your PLAXIS version."
        )

    return found


def filter_interfaces_by_name(interface_list, target_names):
    if not target_names:
        return interface_list

    target_names_lower = [t.lower() for t in target_names]
    filtered = []

    for i, obj in enumerate(interface_list, start=1):
        nm = safe_name(obj, f"Interface_{i}").lower()
        if any(t in nm for t in target_names_lower):
            filtered.append(obj)

    return filtered


def get_interface_result_family(g_o):
    """
    Finds the most likely ResultTypes family for interfaces.
    """
    preferred = [
        "Interface",
        "Interfaces",
        "InterfaceElement",
        "InterfaceElements",
    ]

    for p in preferred:
        if hasattr(g_o.ResultTypes, p):
            return getattr(g_o.ResultTypes, p), p

    names = dir(g_o.ResultTypes)
    for n in names:
        if "interface" in n.lower():
            return getattr(g_o.ResultTypes, n), n

    raise Exception("Could not find an interface ResultTypes family.")


def find_exact_or_suffix(result_family, candidate_names):
    for n in candidate_names:
        if hasattr(result_family, n):
            return getattr(result_family, n), n
    return None, None


def find_result_type_by_keywords(result_family, must_have, optional_groups=None):
    """
    Finds a result type name in result_family by keyword matching.
    must_have: list[str] all must appear
    optional_groups: list[list[str]] at least one from each inner list must appear
    """
    names = dir(result_family)

    for n in names:
        ln = n.lower()

        if not all(k.lower() in ln for k in must_have):
            continue

        ok = True
        if optional_groups:
            for grp in optional_groups:
                if not any(g.lower() in ln for g in grp):
                    ok = False
                    break

        if ok:
            return getattr(result_family, n), n

    return None, None


def get_result_with_locations(g_o, obj, phase, result_type, locations=("stress point", "stresspoint", "node")):
    """
    Tries several result locations because PLAXIS versions may differ.
    """
    last_err = None
    for loc in locations:
        try:
            vals = g_o.getresults(obj, phase, result_type, loc)
            if vals is not None and len(vals) > 0:
                return list(vals), loc
        except Exception as e:
            last_err = e

    if last_err:
        raise last_err

    return [], None


def project_points_to_2d(points3d):
    """
    Best-fit plane projection using SVD/PCA.
    Returns 2D coordinates and centroid.
    """
    pts = np.asarray(points3d, dtype=float)
    centroid = pts.mean(axis=0)
    centered = pts - centroid

    _, _, vh = np.linalg.svd(centered, full_matrices=False)
    e1 = vh[0]
    e2 = vh[1]

    uv = np.column_stack((centered @ e1, centered @ e2))
    return uv, centroid


def triangle_area_3d(p1, p2, p3):
    return 0.5 * np.linalg.norm(np.cross(p2 - p1, p3 - p1))


def unique_points_with_sigma(x, y, z, sigma, decimals=8):
    """
    Removes duplicate points by rounded coordinates.
    Keeps only numeric rows.
    If duplicates exist, sigma is averaged.
    """
    bucket = {}

    for xi, yi, zi, si in zip(x, y, z, sigma):
        xf = safe_float(xi)
        yf = safe_float(yi)
        zf = safe_float(zi)
        sf = safe_float(si)

        if xf is None or yf is None or zf is None or sf is None:
            continue

        key = (
            round(xf, decimals),
            round(yf, decimals),
            round(zf, decimals)
        )

        if key not in bucket:
            bucket[key] = [xf, yf, zf, [sf]]
        else:
            bucket[key][3].append(sf)

    pts = []
    sig = []

    for _, (xf, yf, zf, svals) in bucket.items():
        pts.append([xf, yf, zf])
        sig.append(sum(svals) / len(svals))

    return np.array(pts, dtype=float), np.array(sig, dtype=float)


def integrate_sigma_over_surface(points3d, sigma_vals):
    """
    Approximate surface integral:
        F = ∫ sigma_n dA
    using triangulation of projected points onto best-fit plane.

    Returns:
        area_total, resultant_force
    """
    import matplotlib.tri as mtri

    pts = np.asarray(points3d, dtype=float)
    sigma = np.asarray(sigma_vals, dtype=float)

    if len(pts) < 3:
        return 0.0, 0.0

    uv, _ = project_points_to_2d(pts)

    tri = mtri.Triangulation(uv[:, 0], uv[:, 1])

    area_total = 0.0
    force_total = 0.0

    for tri_idx in tri.triangles:
        i, j, k = tri_idx
        p1, p2, p3 = pts[i], pts[j], pts[k]

        area = triangle_area_3d(p1, p2, p3)
        if area <= 0:
            continue

        sigma_avg = (sigma[i] + sigma[j] + sigma[k]) / 3.0

        area_total += area
        force_total += sigma_avg * area

    return area_total, force_total


# ============================================================
# MAIN
# ============================================================
phase = get_last_phase(g_o)

interfaces = get_all_interfaces(g_o)
interfaces = filter_interfaces_by_name(interfaces, TARGET_INTERFACE_NAMES)

if not interfaces:
    raise Exception("No interfaces matched TARGET_INTERFACE_NAMES.")

result_family, result_family_name = get_interface_result_family(g_o)

print("Using ResultTypes family:", result_family_name)
print("\nAvailable Interface result types:")
for n in dir(result_family):
    print("   ", n)

# ------------------------------------------------------------
# coordinate result types: first try exact/common names
# ------------------------------------------------------------
rt_x, rt_x_name = find_exact_or_suffix(result_family, [
    "X", "CoordX", "CoordinateX", "InterfaceX"
])
rt_y, rt_y_name = find_exact_or_suffix(result_family, [
    "Y", "CoordY", "CoordinateY", "InterfaceY"
])
rt_z, rt_z_name = find_exact_or_suffix(result_family, [
    "Z", "CoordZ", "CoordinateZ", "InterfaceZ"
])

# fallback to keyword search if exact names not found
if rt_x is None:
    rt_x, rt_x_name = find_result_type_by_keywords(result_family, ["x"])
if rt_y is None:
    rt_y, rt_y_name = find_result_type_by_keywords(result_family, ["y"])
if rt_z is None:
    rt_z, rt_z_name = find_result_type_by_keywords(result_family, ["z"])

if rt_x is None or rt_y is None or rt_z is None:
    raise Exception(
        "Could not find interface coordinate result types X/Y/Z. "
        "Check the printed Interface result types above."
    )

print("\nUsing coordinate result types:")
print("   X ->", rt_x_name)
print("   Y ->", rt_y_name)
print("   Z ->", rt_z_name)

# ------------------------------------------------------------
# effective normal stress result type
# ------------------------------------------------------------
sigma_candidates = [
    (["interface", "effective", "normal", "stress"], None),
    (["effective", "normal", "stress"], None),
    (["sigma", "n"], [["eff", "effective"]]),
    (["sigman"], [["eff", "effective"]]),
    (["normal"], [["stress"], ["eff", "effective"]]),
]

rt_sigma = None
rt_sigma_name = None

for must_have, optional_groups in sigma_candidates:
    rt_sigma, rt_sigma_name = find_result_type_by_keywords(result_family, must_have, optional_groups)
    if rt_sigma is not None:
        break

if rt_sigma is None:
    raise Exception(
        "Could not auto-detect effective normal stress result type for interface. "
        "Check the printed Interface result types above."
    )

print("Using sigma_n result type:", rt_sigma_name)

# ============================================================
# EXCEL OUTPUT
# ============================================================
wb = Workbook()

ws_sum = wb.active
ws_sum.title = "Summary"
ws_sum.append([
    "InterfaceName",
    "Phase",
    "ResultType",
    "LocationType",
    "IntegratedArea",
    "ResultantForce",
    "AverageSigma_n_eff",
])

ws_raw = wb.create_sheet("Raw_Data")
ws_raw.append([
    "InterfaceName",
    "PointNo",
    "X",
    "Y",
    "Z",
    "Sigma_n_eff",
])

# ============================================================
# PROCESS EACH INTERFACE
# ============================================================
processed = 0

for idx, interface_obj in enumerate(interfaces, start=1):
    interface_name = safe_name(interface_obj, f"Interface_{idx}")

    try:
        x_vals, loc_x = get_result_with_locations(g_o, interface_obj, phase, rt_x)
        y_vals, loc_y = get_result_with_locations(g_o, interface_obj, phase, rt_y)
        z_vals, loc_z = get_result_with_locations(g_o, interface_obj, phase, rt_z)
        s_vals, loc_s = get_result_with_locations(g_o, interface_obj, phase, rt_sigma)
    except Exception as e:
        print(f"\nSkipping {interface_name}: failed to read results -> {e}")
        continue

    print(f"\nInterface: {interface_name}")
    print("Sample X:", x_vals[:10])
    print("Sample Y:", y_vals[:10])
    print("Sample Z:", z_vals[:10])
    print("Sample sigma:", s_vals[:10])

    n = min(len(x_vals), len(y_vals), len(z_vals), len(s_vals))
    if n < 3:
        print(f"Skipping {interface_name}: not enough points.")
        continue

    x_vals = x_vals[:n]
    y_vals = y_vals[:n]
    z_vals = z_vals[:n]
    s_vals = s_vals[:n]

    pts, sig = unique_points_with_sigma(x_vals, y_vals, z_vals, s_vals)

    if len(pts) < 3:
        print(f"Skipping {interface_name}: not enough valid numeric points after filtering.")
        continue

    area_total, force_total = integrate_sigma_over_surface(pts, sig)

    if area_total > 0:
        avg_sigma = force_total / area_total
    else:
        avg_sigma = None

    ws_sum.append([
        interface_name,
        safe_name(phase, "Last phase"),
        rt_sigma_name,
        loc_s,
        area_total,
        force_total,
        avg_sigma,
    ])

    for pno, (p, s) in enumerate(zip(pts, sig), start=1):
        ws_raw.append([
            interface_name,
            pno,
            p[0],
            p[1],
            p[2],
            s,
        ])

    processed += 1

    if avg_sigma is None:
        print(
            f"Processed: {interface_name} | "
            f"Area = {area_total:.6f} | "
            f"Resultant = {force_total:.6f} | "
            f"Average = None"
        )
    else:
        print(
            f"Processed: {interface_name} | "
            f"Area = {area_total:.6f} | "
            f"Resultant = {force_total:.6f} | "
            f"Average = {avg_sigma:.6f}"
        )

# ============================================================
# FORMATTING + SAVE
# ============================================================
for ws in [ws_sum, ws_raw]:
    ws.freeze_panes = "A2"

for col in ["A", "B", "C", "D", "E", "F", "G"]:
    ws_sum.column_dimensions[col].width = 22

for col in ["A", "B", "C", "D", "E", "F"]:
    ws_raw.column_dimensions[col].width = 18

wb.save(OUT_FILE)

print("\nDONE")
print("Excel file:", OUT_FILE)
print("Interfaces processed:", processed)
print("Phase:", safe_name(phase, "Last phase"))