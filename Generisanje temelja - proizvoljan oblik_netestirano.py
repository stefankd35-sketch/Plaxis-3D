import sys
sys.path.append(r"C:\ProgramData\Seequent\PLAXIS Python Distribution V3\python\Lib\site-packages")
import ezdxf
import math
from pathlib import Path
from plxscripting.easy import *

# ============================================================
# USER SETTINGS - UPDATED TO MATCH YOUR DXF SCAN
# ============================================================
DXF_FILE = Path(r"C:\Users\jelen\OneDrive\Desktop\pyPlaxis3D\layout.dxf")

# Matches your DXF scan exactly
PILE_LAYER = "PILES"
SLAB_LAYER = "SLAB_BOUNDARY"

# PLAXIS Materials (Ensure these exist in your project!)
PILE_MAT_NAME = "Pile_Material_01"    
PLATE_MAT_NAME = "Raft_Concrete" 

# Vertical Settings
Z_TOP_PILE = -5.0
Z_BOT_PILE = -15.0

# Plaxis Connection
PLAXIS_HOST, PLAXIS_PORT, PLAXIS_PASSWORD = "localhost", 10000, "12345"

# ============================================================
# DATA EXTRACTION
# ============================================================
def get_complex_data(dxf_path):
    try:
        doc = ezdxf.readfile(dxf_path)
        msp = doc.modelspace()
        
        # 1. PILES
        piles = []
        # Query using the exact layer name from your scan
        for e in msp.query(f'*[layer=="{PILE_LAYER}"]'):
            if e.dxftype() == 'POINT':
                piles.append((e.dxf.location.x, e.dxf.location.y))
            elif e.dxftype() == 'CIRCLE':
                piles.append((e.dxf.center.x, e.dxf.center.y))
        
        # 2. SLAB (Lines)
        entities = msp.query(f'*[layer=="{SLAB_LAYER}"]')
        slab_pts = []
        for e in entities:
            if e.dxftype() == 'LINE':
                # Rounding to 3 decimals to avoid CAD precision errors
                p1 = (round(e.dxf.start.x, 3), round(e.dxf.start.y, 3), round(e.dxf.start.z, 3))
                p2 = (round(e.dxf.end.x, 3), round(e.dxf.end.y, 3), round(e.dxf.end.z, 3))
                if p1 not in slab_pts: slab_pts.append(p1)
                if p2 not in slab_pts: slab_pts.append(p2)
        
        return piles, slab_pts
    except Exception as e:
        print(f"DXF Error: {e}")
        return [], []

# ============================================================
# EXECUTION
# ============================================================
try:
    s_i, g_i = new_server(PLAXIS_HOST, PLAXIS_PORT, password=PLAXIS_PASSWORD)
    g_i.gotostructures()

    piles, slab_pts = get_complex_data(DXF_FILE)

    # --- PART 1: SLAB & 1:2 SLOPED EXCAVATION ---
    if len(slab_pts) >= 3:
        print(f"Creating Slab with {len(slab_pts)} vertices...")
        z_bot = slab_pts[0][2]
        offset_dist = 2 * abs(z_bot)
        
        # Centroid for outward offset direction
        avg_x = sum(p[0] for p in slab_pts) / len(slab_pts)
        avg_y = sum(p[1] for p in slab_pts) / len(slab_pts)

        # Bottom Plate
        bottom_surf = g_i.surface(*slab_pts)
        plate = g_i.plate(bottom_surf)
        
        plate_mat = next((m for m in g_i.Materials[:] if m.Name.value == PLATE_MAT_NAME), None)
        if plate_mat: g_i.setmaterial(plate, plate_mat)

        # Top Vertices (Z=0) and Slopes
        top_pts = []
        for px, py, pz in slab_pts:
            dx, dy = px - avg_x, py - avg_y
            dist = math.sqrt(dx**2 + dy**2)
            ox = (dx / dist) * offset_dist if dist != 0 else 0
            oy = (dy / dist) * offset_dist if dist != 0 else 0
            top_pts.append((px + ox, py + oy, 0))

        g_i.surface(*top_pts)
        for i in range(len(slab_pts)):
            b1, b2 = slab_pts[i], slab_pts[(i + 1) % len(slab_pts)]
            t1, t2 = top_pts[i], top_pts[(i + 1) % len(slab_pts)]
            g_i.surface(b1, b2, t2, t1)
        print("Slab and sloped walls created.")

    # --- PART 2: PILES ---
    if piles:
        print(f"Found {len(piles)} piles. Creating...")
        pile_mat = next((m for m in g_i.Materials[:] if m.Name.value == PILE_MAT_NAME), None)
        for i, (px, py) in enumerate(piles):
            res = g_i.embeddedbeam((px, py, Z_TOP_PILE), (px, py, Z_BOT_PILE))
            actual_beam = [obj for obj in res if obj.TypeName.value == "EmbeddedBeam"][0]
            if pile_mat: g_i.setmaterial(actual_beam, pile_mat)
        print("Piles created.")

except Exception as e:
    print(f"Error during execution: {e}")

print("--- Process Finished ---")