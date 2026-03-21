import sys
sys.path.append(r"C:\ProgramData\Seequent\PLAXIS Python Distribution V3\python\Lib\site-packages")

import ezdxf
from pathlib import Path
from collections import defaultdict
from plxscripting.easy import *

# ============================================================
# USER SETTINGS
# ============================================================
DXF_FILE = Path(r"C:\Users\jelen\OneDrive\Desktop\pyPlaxis3D\layout.dxf")

PILE_LAYER = "PILES"
BOREHOLE_LAYER = "BOREHOLES"

PILE_MAT_NAME = "Pile_Material_01"
PLATE_MAT_NAME = "Raft_Concrete"

# PLAXIS connection
PLAXIS_HOST = "localhost"
PLAXIS_PORT = 10000
PLAXIS_PASSWORD = "12345"

# Tolerances
Z_TOL = 1e-6
XY_ROUND = 3

# ============================================================
# HELPERS
# ============================================================
def round_pt(x, y, z=None):
    if z is None:
        return (round(x, XY_ROUND), round(y, XY_ROUND))
    return (round(x, XY_ROUND), round(y, XY_ROUND), round(z, XY_ROUND))


def is_zero(val, tol=Z_TOL):
    return abs(val) <= tol


def classify_line(p1, p2):
    z1, z2 = p1[2], p2[2]

    if z1 < -Z_TOL and z2 < -Z_TOL:
        return "plate"
    if is_zero(z1) and is_zero(z2):
        return "surface"
    return None


def xy_dist2(p1, p2):
    return (p1[0] - p2[0])**2 + (p1[1] - p2[1])**2


def rotate_list(lst, shift):
    shift = shift % len(lst)
    return lst[shift:] + lst[:shift]


def build_ordered_polygon_from_segments(segments):
    if not segments:
        return []

    adjacency = defaultdict(list)
    point_map = {}

    for p1, p2 in segments:
        k1 = (p1[0], p1[1])
        k2 = (p2[0], p2[1])

        adjacency[k1].append(k2)
        adjacency[k2].append(k1)

        point_map[k1] = p1
        point_map[k2] = p2

    # check closed-loop logic
    bad_nodes = [k for k, v in adjacency.items() if len(v) != 2]
    if bad_nodes:
        raise RuntimeError(
            "DXF boundary is not a single clean closed polygon. "
            f"Found vertices with degree != 2: {bad_nodes[:10]}"
        )

    start = next(iter(adjacency.keys()))
    ordered_keys = [start]
    prev_key = None
    current_key = start

    while True:
        neighbors = adjacency[current_key]

        if prev_key is None:
            next_key = neighbors[0]
        else:
            candidates = [n for n in neighbors if n != prev_key]
            if not candidates:
                break
            next_key = candidates[0]

        if next_key == start:
            break

        ordered_keys.append(next_key)
        prev_key = current_key
        current_key = next_key

        if len(ordered_keys) > len(point_map) + 5:
            raise RuntimeError("Failed to reconstruct closed polygon from DXF lines.")

    return [point_map[k] for k in ordered_keys]


def align_surface_to_plate(plate_pts, surface_pts):
    """
    Reorder surface_pts so they correspond to plate_pts:
    - same start corner
    - same traversal direction
    Uses XY proximity minimization.
    """
    if len(plate_pts) != len(surface_pts):
        raise RuntimeError(
            f"Plate boundary has {len(plate_pts)} vertices, "
            f"but top surface boundary has {len(surface_pts)} vertices."
        )

    best_pts = None
    best_score = None
    n = len(plate_pts)

    candidates = [
        surface_pts[:],
        list(reversed(surface_pts))
    ]

    for candidate in candidates:
        for shift in range(n):
            rotated = rotate_list(candidate, shift)
            score = sum(xy_dist2(plate_pts[i], rotated[i]) for i in range(n))
            if best_score is None or score < best_score:
                best_score = score
                best_pts = rotated

    return best_pts


def get_dxf_data(dxf_path):
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    pile_points = []
    borehole_points = []
    plate_segments = []
    surface_segments = []

    for e in msp:
        layer = e.dxf.layer

        if layer in (PILE_LAYER, BOREHOLE_LAYER):
            if e.dxftype() == "POINT":
                pt = round_pt(e.dxf.location.x, e.dxf.location.y)
            elif e.dxftype() == "CIRCLE":
                pt = round_pt(e.dxf.center.x, e.dxf.center.y)
            else:
                continue

            if layer == PILE_LAYER:
                pile_points.append(pt)
            else:
                borehole_points.append(pt)

        elif e.dxftype() == "LINE":
            p1 = round_pt(e.dxf.start.x, e.dxf.start.y, e.dxf.start.z)
            p2 = round_pt(e.dxf.end.x, e.dxf.end.y, e.dxf.end.z)

            kind = classify_line(p1, p2)
            if kind == "plate":
                plate_segments.append((p1, p2))
            elif kind == "surface":
                surface_segments.append((p1, p2))

    plate_pts = build_ordered_polygon_from_segments(plate_segments)
    surface_pts = build_ordered_polygon_from_segments(surface_segments)

    return pile_points, borehole_points, plate_pts, surface_pts


def get_material_by_name(g_i, mat_name):
    return next((m for m in g_i.Materials[:] if m.Name.value == mat_name), None)


def create_borehole(g_i, x, y):
    errors = []

    for args in [
        (x, y),
        ((x, y),),
        ((x, y, 0.0),),
        (x,)
    ]:
        try:
            return g_i.borehole(*args)
        except Exception as e:
            errors.append(f"borehole{args}: {e}")

    raise RuntimeError(
        f"Could not create borehole at ({x}, {y}).\n" + "\n".join(errors)
    )


def get_pile_length():
    while True:
        try:
            val = float(input("Enter pile length (positive value, in meters): "))
            if val <= 0:
                print("Pile length must be a positive number.")
                continue
            return val
        except ValueError:
            print("Invalid input. Please enter a numeric value.")


def create_sloped_side_surfaces(g_i, plate_pts, surface_pts):
    """
    Create sloped side surfaces between plate polygon and top surface polygon.
    Automatically aligns top boundary to the plate boundary.
    """
    if len(plate_pts) < 3 or len(surface_pts) < 3:
        raise RuntimeError("Need at least 3 points in both plate and surface boundaries.")

    if len(plate_pts) != len(surface_pts):
        raise RuntimeError(
            f"Plate boundary has {len(plate_pts)} vertices, "
            f"but top surface boundary has {len(surface_pts)} vertices. "
            "They must have the same number of vertices."
        )

    aligned_surface_pts = align_surface_to_plate(plate_pts, surface_pts)

    print("Plate to top point matching:")
    for i, (bp, tp) in enumerate(zip(plate_pts, aligned_surface_pts), start=1):
        print(f"  {i}: plate {bp} -> top {tp}")

    side_surfaces = []
    n = len(plate_pts)

    for i in range(n):
        b1 = plate_pts[i]
        b2 = plate_pts[(i + 1) % n]
        t2 = aligned_surface_pts[(i + 1) % n]
        t1 = aligned_surface_pts[i]

        side = g_i.surface(b1, b2, t2, t1)
        side_surfaces.append(side)

    return side_surfaces


# ============================================================
# EXECUTION
# ============================================================
try:
    s_i, g_i = new_server(PLAXIS_HOST, PLAXIS_PORT, password=PLAXIS_PASSWORD)
    g_i.gotostructures()

    piles, boreholes, plate_pts, surface_pts = get_dxf_data(DXF_FILE)

    print("DXF read complete:")
    print(f"  Plate boundary points   : {len(plate_pts)}")
    print(f"  Surface boundary points : {len(surface_pts)}")
    print(f"  Piles                   : {len(piles)}")
    print(f"  Boreholes               : {len(boreholes)}")

    plate_z = None

    # --------------------------------------------------------
    # PART 1: Plate from lines with z < 0
    # --------------------------------------------------------
    if len(plate_pts) >= 3:
        plate_z = plate_pts[0][2]

        for p in plate_pts:
            if abs(p[2] - plate_z) > Z_TOL:
                raise RuntimeError("Plate boundary points do not have the same Z elevation.")

        print(f"Creating plate surface with {len(plate_pts)} vertices at Z = {plate_z}...")
        plate_surf = g_i.surface(*plate_pts)
        plate_obj = g_i.plate(plate_surf)

        plate_mat = get_material_by_name(g_i, PLATE_MAT_NAME)
        if plate_mat:
            g_i.setmaterial(plate_obj, plate_mat)
            print(f'Assigned plate material: "{PLATE_MAT_NAME}"')
        else:
            print(f'Warning: Plate material "{PLATE_MAT_NAME}" not found.')
    else:
        print("Warning: Not enough z<0 line data to create plate surface.")

    # --------------------------------------------------------
    # PART 2: Ground/surface from lines with z = 0
    # --------------------------------------------------------
    if len(surface_pts) >= 3:
        print(f"Creating ground surface with {len(surface_pts)} vertices...")
        g_i.surface(*surface_pts)
    else:
        print("Warning: Not enough z=0 line data to create top surface.")

    # --------------------------------------------------------
    # PART 2B: Sloped side surfaces between plate and top
    # --------------------------------------------------------
    if len(plate_pts) >= 3 and len(surface_pts) >= 3:
        print("Creating sloped side surfaces...")
        side_surfaces = create_sloped_side_surfaces(g_i, plate_pts, surface_pts)
        print(f"Created {len(side_surfaces)} sloped side surfaces.")

    # --------------------------------------------------------
    # PART 3: Piles from layer PILES
    # Top = plate level
    # Bottom = plate level - pile length
    # --------------------------------------------------------
    if piles:
        if plate_z is None:
            raise RuntimeError("Cannot create piles because plate elevation could not be determined.")

        pile_length = get_pile_length()

        print(f"Creating {len(piles)} piles...")
        pile_mat = get_material_by_name(g_i, PILE_MAT_NAME)

        z_top_pile = plate_z
        z_bot_pile = plate_z - pile_length

        print(f"Pile top level    = {z_top_pile}")
        print(f"Pile bottom level = {z_bot_pile}")
        print(f"Pile length       = {pile_length}")

        for i, (px, py) in enumerate(piles, start=1):
            try:
                res = g_i.embeddedbeam((px, py, z_top_pile), (px, py, z_bot_pile))
                actual_beam = next(obj for obj in res if obj.TypeName.value == "EmbeddedBeam")

                if pile_mat:
                    g_i.setmaterial(actual_beam, pile_mat)

                print(f"  Pile {i}: ({px}, {py})")
            except Exception as e:
                print(f"  Failed to create pile {i} at ({px}, {py}): {e}")

        if not pile_mat:
            print(f'Warning: Pile material "{PILE_MAT_NAME}" not found.')
    else:
        print("No pile points found on layer PILES.")

    # --------------------------------------------------------
    # PART 4: Boreholes from layer BOREHOLES
    # --------------------------------------------------------
    if boreholes:
        print(f"Creating {len(boreholes)} boreholes...")
        for i, (bx, by) in enumerate(boreholes, start=1):
            try:
                create_borehole(g_i, bx, by)
                print(f"  Borehole {i}: ({bx}, {by})")
            except Exception as e:
                print(f"  Failed to create borehole {i} at ({bx}, {by}): {e}")
    else:
        print("No borehole points found on layer BOREHOLES.")

except Exception as e:
    print(f"Error during execution: {e}")

print("--- Process Finished ---")