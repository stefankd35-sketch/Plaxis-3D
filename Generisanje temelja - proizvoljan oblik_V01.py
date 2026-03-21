import sys
sys.path.append(r"C:\ProgramData\Seequent\PLAXIS Python Distribution V3\python\Lib\site-packages")
import ezdxf
import math
from pathlib import Path
from plxscripting.easy import *

# ============================================================
# USER SETTINGS
# ============================================================
DXF_FILE = Path(r"C:\Users\jelen\OneDrive\Desktop\pyPlaxis3D\layout.dxf")
PILE_LAYER = "PILES"
SLAB_LAYER = "SLAB_BOUNDARY"

PILE_MAT_NAME = "Concrete_Pile_Material"    
PLATE_MAT_NAME = "Raft_Material" 

# ============================================================
# EXECUTION
# ============================================================
try:
    s_i, g_i = new_server("localhost", 10000, password="12345")
    g_i.gotostructures()

    # Load DXF Data
    doc = ezdxf.readfile(DXF_FILE)
    msp = doc.modelspace()
    
    # Extract Slab Vertices
    entities = msp.query(f'*[layer=="{SLAB_LAYER}"]')
    slab_pts = []
    for e in entities:
        if e.dxftype() == 'LINE':
            p1 = (round(e.dxf.start.x, 3), round(e.dxf.start.y, 3), round(e.dxf.start.z, 3))
            p2 = (round(e.dxf.end.x, 3), round(e.dxf.end.y, 3), round(e.dxf.end.z, 3))
            if p1 not in slab_pts: slab_pts.append(p1)
            if p2 not in slab_pts: slab_pts.append(p2)

    if len(slab_pts) >= 3:
        z_bot = slab_pts[0][2]
        offset_dist = 2 * abs(z_bot) # 1:2 Slope
        
        avg_x = sum(p[0] for p in slab_pts) / len(slab_pts)
        avg_y = sum(p[1] for p in slab_pts) / len(slab_pts)

        # 1. BOTTOM SURFACE & PLATE
        bottom_surf = g_i.surface(*slab_pts)
        plate = g_i.plate(bottom_surf)
        
        # 2. CALCULATE TOP POINTS (Z=0)
        top_pts = []
        for px, py, pz in slab_pts:
            dx, dy = px - avg_x, py - avg_y
            dist = math.sqrt(dx**2 + dy**2)
            ox = (dx / dist) * offset_dist if dist != 0 else 0
            oy = (dy / dist) * offset_dist if dist != 0 else 0
            top_pts.append((px + ox, py + oy, 0))

        # 3. CREATE TOP SURFACE (Z=0)
        top_surf = g_i.surface(*top_pts)

        # 4. CREATE SIDE SLOPE SURFACES
        # We loop through each segment and create a 4-point surface
        side_surfaces = []
        for i in range(len(slab_pts)):
            # Segment endpoints
            b1, b2 = slab_pts[i], slab_pts[(i + 1) % len(slab_pts)]
            t1, t2 = top_pts[i], top_pts[(i + 1) % len(slab_pts)]
            
            # This creates the actual sloped Surface object
            s = g_i.surface(b1, b2, t2, t1)
            side_surfaces.append(s)

        # 5. GROUPING (Optional but helpful)
        g_i.group(side_surfaces, "Excavation_Slopes")
        print(f"Successfully transformed {len(side_surfaces)} side polygons into Surfaces.")

except Exception as e:
    print(f"Error: {e}")

print("--- Process Finished ---")