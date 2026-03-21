import sys
sys.path.append(r"C:\ProgramData\Seequent\PLAXIS Python Distribution V3\python\Lib\site-packages")

import ezdxf
from pathlib import Path
from collections import defaultdict
from plxscripting.easy import *

# ============================================================
# USER SETTINGS
# ============================================================
DXF_FILE = Path(r"C:\Users\jelen\OneDrive\Desktop\Plaxis 3D - Podloge\layout.dxf")

PILE_LAYER = "PILES"
BOREHOLE_LAYER = "BOREHOLES"
SLAB_BOUNDARY_LAYER = "SLAB_BOUNDARY"
POINT_LINE_LOAD_LAYER = "POINT-LINE LOADS"
SURFACE_LOAD_LAYER = "SURFACE LOADS"

PILE_MAT_NAME = "Pile_Material_01"
PLATE_MAT_NAME = "Raft_Concrete"

DEFAULT_POINT_LOAD_Z = 0.0
DEFAULT_LINE_LOAD_Z = 0.0
DEFAULT_SURFACE_LOAD_Z = 0.0

PLAXIS_HOST = "localhost"
PLAXIS_PORT = 10000
PLAXIS_PASSWORD = "12345"

Z_TOL = 1e-6
XY_ROUND = 3
KEY_DIGITS = 3

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


def key2d(pt, ndigits=KEY_DIGITS):
    return (round(pt[0], ndigits), round(pt[1], ndigits))


def normalize_layer_name(name):
    return str(name).strip().upper()


def layer_matches(entity_layer, target_layer):
    return normalize_layer_name(entity_layer) == normalize_layer_name(target_layer)


def print_available_layers(doc):
    try:
        print("Available DXF layers:")
        for layer in doc.layers:
            print(f"  - '{layer.dxf.name}'")
    except Exception as e:
        print(f"Could not list layers: {e}")


def get_lwpolyline_points(e):
    pts = []
    for p in e.get_points():
        x = p[0]
        y = p[1]
        pts.append(round_pt(x, y))
    return pts


def get_polyline_points(e):
    pts = []
    try:
        for v in e.vertices:
            pts.append(round_pt(v.dxf.location.x, v.dxf.location.y))
    except Exception:
        try:
            for v in e.points():
                pts.append(round_pt(v[0], v[1]))
        except Exception:
            pass
    return pts


def polyline_is_closed(e):
    try:
        return bool(e.closed)
    except Exception:
        try:
            return bool(e.is_closed)
        except Exception:
            return False


def points_to_segments_2d(pts):
    segs = []
    if len(pts) < 2:
        return segs
    for i in range(len(pts) - 1):
        segs.append((pts[i], pts[i + 1]))
    return segs


def closed_points_to_segments_3d(pts3d):
    segs = []
    if len(pts3d) < 3:
        return segs
    for i in range(len(pts3d)):
        segs.append((pts3d[i], pts3d[(i + 1) % len(pts3d)]))
    return segs


def segment_groups(segments):
    groups = []

    for seg in segments:
        p1, p2 = seg
        k1 = key2d(p1)
        k2 = key2d(p2)

        matched = []
        for i, grp in enumerate(groups):
            if k1 in grp["keys"] or k2 in grp["keys"]:
                matched.append(i)

        if not matched:
            groups.append({
                "segments": [seg],
                "keys": {k1, k2}
            })
        else:
            first = matched[0]
            groups[first]["segments"].append(seg)
            groups[first]["keys"].update([k1, k2])

            for j in reversed(matched[1:]):
                groups[first]["segments"].extend(groups[j]["segments"])
                groups[first]["keys"].update(groups[j]["keys"])
                groups.pop(j)

    return [g["segments"] for g in groups]


def build_ordered_polygon_from_segments_tolerant(segments):
    if not segments:
        return [], False, []

    adjacency = defaultdict(list)
    point_map = {}

    for p1, p2 in segments:
        k1 = key2d(p1)
        k2 = key2d(p2)

        adjacency[k1].append(k2)
        adjacency[k2].append(k1)

        point_map[k1] = p1
        point_map[k2] = p2

    bad_nodes = [k for k, v in adjacency.items() if len(v) != 2]
    is_closed = len(bad_nodes) == 0

    if not is_closed:
        return [], False, bad_nodes

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
                return [], False, ["Traversal failed"]
            next_key = candidates[0]

        if next_key == start:
            break

        ordered_keys.append(next_key)
        prev_key = current_key
        current_key = next_key

        if len(ordered_keys) > len(point_map) + 5:
            return [], False, ["Traversal failed"]

    return [point_map[k] for k in ordered_keys], True, []


def split_closed_loops(segments, label="boundary"):
    if not segments:
        return []

    loops = []
    groups = segment_groups(segments)

    for i, grp in enumerate(groups, start=1):
        polygon, is_closed, bad_nodes = build_ordered_polygon_from_segments_tolerant(grp)

        if is_closed:
            loops.append(polygon)
        else:
            print(f"Warning: Skipping open {label} group {i}. Bad vertices: {bad_nodes[:10]}")

    return loops


def align_surface_to_plate(plate_pts, surface_pts):
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


def safe_set_load_z(obj, value):
    if value is None:
        return False

    candidate_props = [
        "Fz", "qz", "pz",
        "sigmaz",
        "z", "z_start", "z_end",
        "qy", "qx"
    ]

    for prop in candidate_props:
        try:
            getattr(obj, prop).set(value)
            return True
        except Exception:
            pass

    return False


def create_point_load(g_i, x, y, z, load_value_z=None):
    errors = []

    for args in [
        ((x, y, z),),
        (x, y, z),
    ]:
        try:
            obj = g_i.pointload(*args)
            safe_set_load_z(obj, load_value_z)
            return obj
        except Exception as e:
            errors.append(f"pointload{args}: {e}")

    raise RuntimeError(
        f"Could not create point load at ({x}, {y}, {z}).\n" + "\n".join(errors)
    )


def create_line_load(g_i, p1, p2, load_value_z=None):
    errors = []

    for args in [
        (p1, p2),
        (*p1, *p2),
    ]:
        try:
            obj = g_i.lineload(*args)
            safe_set_load_z(obj, load_value_z)
            return obj
        except Exception as e:
            errors.append(f"lineload{args}: {e}")

    raise RuntimeError(
        f"Could not create line load from {p1} to {p2}.\n" + "\n".join(errors)
    )


def create_surface_load(g_i, pts3d, load_value_z=None):
    errors = []

    try:
        surf = g_i.surface(*pts3d)
    except Exception as e:
        raise RuntimeError(f"Could not create support surface for surface load: {e}")

    for args in [
        (surf,),
        tuple(pts3d),
    ]:
        try:
            obj = g_i.surfaceload(*args)
            safe_set_load_z(obj, load_value_z)
            return obj
        except Exception as e:
            errors.append(f"surfaceload{args}: {e}")

    raise RuntimeError("Could not create surface load.\n" + "\n".join(errors))


def get_dxf_data(dxf_path):
    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()

    print_available_layers(doc)

    pile_points = []
    borehole_points = []
    plate_segments = []
    surface_segments = []

    point_load_points = []
    line_load_lines = []
    surface_load_segments = []
    surface_load_polygons = []

    for e in msp:
        layer = e.dxf.layer
        dtype = e.dxftype()

        # ----------------------------------------------------
        # PILES / BOREHOLES
        # ----------------------------------------------------
        if layer_matches(layer, PILE_LAYER) or layer_matches(layer, BOREHOLE_LAYER):
            if dtype == "POINT":
                pt = round_pt(e.dxf.location.x, e.dxf.location.y)
            elif dtype == "CIRCLE":
                pt = round_pt(e.dxf.center.x, e.dxf.center.y)
            elif dtype == "INSERT":
                pt = round_pt(e.dxf.insert.x, e.dxf.insert.y)
            else:
                continue

            if layer_matches(layer, PILE_LAYER):
                pile_points.append(pt)
            else:
                borehole_points.append(pt)

        # ----------------------------------------------------
        # POINT/LINE LOADS
        # ----------------------------------------------------
        elif layer_matches(layer, POINT_LINE_LOAD_LAYER):
            if dtype == "POINT":
                point_load_points.append(round_pt(e.dxf.location.x, e.dxf.location.y))
            elif dtype == "CIRCLE":
                point_load_points.append(round_pt(e.dxf.center.x, e.dxf.center.y))
            elif dtype == "INSERT":
                point_load_points.append(round_pt(e.dxf.insert.x, e.dxf.insert.y))
            elif dtype == "LINE":
                p1 = round_pt(e.dxf.start.x, e.dxf.start.y)
                p2 = round_pt(e.dxf.end.x, e.dxf.end.y)
                line_load_lines.append((p1, p2))
            elif dtype == "LWPOLYLINE":
                pts = get_lwpolyline_points(e)
                segs = points_to_segments_2d(pts)
                line_load_lines.extend(segs)
            elif dtype == "POLYLINE":
                pts = get_polyline_points(e)
                segs = points_to_segments_2d(pts)
                line_load_lines.extend(segs)

        # ----------------------------------------------------
        # SURFACE LOADS
        # ----------------------------------------------------
        elif layer_matches(layer, SURFACE_LOAD_LAYER):
            if dtype == "LINE":
                p1 = round_pt(e.dxf.start.x, e.dxf.start.y, 0.0)
                p2 = round_pt(e.dxf.end.x, e.dxf.end.y, 0.0)
                surface_load_segments.append((p1, p2))

            elif dtype == "LWPOLYLINE":
                pts2d = get_lwpolyline_points(e)
                if polyline_is_closed(e) and len(pts2d) >= 3:
                    surface_load_polygons.append([(p[0], p[1], 0.0) for p in pts2d])
                else:
                    segs = closed_points_to_segments_3d([(p[0], p[1], 0.0) for p in pts2d]) if polyline_is_closed(e) else points_to_segments_2d(pts2d)
                    for a, b in segs:
                        surface_load_segments.append(((a[0], a[1], 0.0), (b[0], b[1], 0.0)))

            elif dtype == "POLYLINE":
                pts2d = get_polyline_points(e)
                if polyline_is_closed(e) and len(pts2d) >= 3:
                    surface_load_polygons.append([(p[0], p[1], 0.0) for p in pts2d])
                else:
                    segs = closed_points_to_segments_3d([(p[0], p[1], 0.0) for p in pts2d]) if polyline_is_closed(e) else points_to_segments_2d(pts2d)
                    for a, b in segs:
                        surface_load_segments.append(((a[0], a[1], 0.0), (b[0], b[1], 0.0)))

        # ----------------------------------------------------
        # SLAB BOUNDARY ONLY
        # ----------------------------------------------------
        elif layer_matches(layer, SLAB_BOUNDARY_LAYER) and dtype == "LINE":
            p1 = round_pt(e.dxf.start.x, e.dxf.start.y, e.dxf.start.z)
            p2 = round_pt(e.dxf.end.x, e.dxf.end.y, e.dxf.end.z)

            kind = classify_line(p1, p2)
            if kind == "plate":
                plate_segments.append((p1, p2))
            elif kind == "surface":
                surface_segments.append((p1, p2))

    plate_loops = split_closed_loops(plate_segments, label="plate")
    surface_loops = split_closed_loops(surface_segments, label="surface")
    surface_load_loops = split_closed_loops(surface_load_segments, label="surface load")
    surface_load_loops.extend(surface_load_polygons)

    if not plate_loops:
        plate_pts = []
    else:
        if len(plate_loops) > 1:
            print(f"Warning: Found {len(plate_loops)} plate loops on layer {SLAB_BOUNDARY_LAYER}. Using the first one.")
        plate_pts = plate_loops[0]

    if not surface_loops:
        surface_pts = []
    else:
        if len(surface_loops) > 1:
            print(f"Warning: Found {len(surface_loops)} top surface loops on layer {SLAB_BOUNDARY_LAYER}. Using the first one.")
        surface_pts = surface_loops[0]

    return {
        "piles": pile_points,
        "boreholes": borehole_points,
        "plate_pts": plate_pts,
        "surface_pts": surface_pts,
        "point_load_points": point_load_points,
        "line_load_lines": line_load_lines,
        "surface_load_loops": surface_load_loops,
    }


# ============================================================
# EXECUTION
# ============================================================
try:
    s_i, g_i = new_server(PLAXIS_HOST, PLAXIS_PORT, password=PLAXIS_PASSWORD)
    g_i.gotostructures()

    data = get_dxf_data(DXF_FILE)

    piles = data["piles"]
    boreholes = data["boreholes"]
    plate_pts = data["plate_pts"]
    surface_pts = data["surface_pts"]
    point_load_points = data["point_load_points"]
    line_load_lines = data["line_load_lines"]
    surface_load_loops = data["surface_load_loops"]

    print("DXF read complete:")
    print(f"  Plate boundary points   : {len(plate_pts)}")
    print(f"  Surface boundary points : {len(surface_pts)}")
    print(f"  Piles                   : {len(piles)}")
    print(f"  Boreholes               : {len(boreholes)}")
    print(f"  Point loads             : {len(point_load_points)}")
    print(f"  Line loads              : {len(line_load_lines)}")
    print(f"  Surface load areas      : {len(surface_load_loops)}")

    plate_z = None

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
        print("Warning: Not enough z<0 line data on layer SLAB BOUNDARY to create plate surface.")

    if len(surface_pts) >= 3:
        print(f"Creating ground surface with {len(surface_pts)} vertices...")
        g_i.surface(*surface_pts)
    else:
        print("Warning: Not enough z=0 line data on layer SLAB BOUNDARY to create top surface.")

    if len(plate_pts) >= 3 and len(surface_pts) >= 3:
        print("Creating sloped side surfaces...")
        side_surfaces = create_sloped_side_surfaces(g_i, plate_pts, surface_pts)
        print(f"Created {len(side_surfaces)} sloped side surfaces.")

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
    else:
        print("No pile points found on layer PILES.")

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

    if point_load_points:
        if plate_z is None:
            raise RuntimeError("Cannot create point loads because plate elevation could not be determined.")

        print(f"Creating {len(point_load_points)} point loads at plate elevation...")
        for i, (x, y) in enumerate(point_load_points, start=1):
            try:
                create_point_load(g_i, x, y, plate_z, DEFAULT_POINT_LOAD_Z)
                print(f"  Point load {i}: ({x}, {y}, {plate_z})")
            except Exception as e:
                print(f"  Failed to create point load {i} at ({x}, {y}, {plate_z}): {e}")
    else:
        print("No point loads found on layer POINT/LINE LOADS.")

    if line_load_lines:
        if plate_z is None:
            raise RuntimeError("Cannot create line loads because plate elevation could not be determined.")

        print(f"Creating {len(line_load_lines)} line loads at plate elevation...")
        for i, (p1_xy, p2_xy) in enumerate(line_load_lines, start=1):
            p1 = (p1_xy[0], p1_xy[1], plate_z)
            p2 = (p2_xy[0], p2_xy[1], plate_z)
            try:
                create_line_load(g_i, p1, p2, DEFAULT_LINE_LOAD_Z)
                print(f"  Line load {i}: {p1} -> {p2}")
            except Exception as e:
                print(f"  Failed to create line load {i}: {e}")
    else:
        print("No line loads found on layer POINT/LINE LOADS.")

    if surface_load_loops:
        if plate_z is None:
            raise RuntimeError("Cannot create surface loads because plate elevation could not be determined.")

        print(f"Creating {len(surface_load_loops)} surface load area(s) at plate elevation...")
        for i, loop in enumerate(surface_load_loops, start=1):
            try:
                loop_at_plate = [(p[0], p[1], plate_z) for p in loop]
                create_surface_load(g_i, loop_at_plate, DEFAULT_SURFACE_LOAD_Z)
                print(f"  Surface load {i}: {len(loop_at_plate)} vertices")
            except Exception as e:
                print(f"  Failed to create surface load {i}: {e}")
    else:
        print("No closed surface loads found on layer SURFACE LOADS.")

except Exception as e:
    print(f"Error during execution: {e}")

print("--- Process Finished ---")