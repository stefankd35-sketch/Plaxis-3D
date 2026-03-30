import re
import sys
sys.path.append(r"C:\ProgramData\Seequent\PLAXIS Python Distribution V3\python\Lib\site-packages")

from plxscripting.easy import *
from openpyxl import Workbook
import matplotlib.pyplot as plt

# =========================
# SETTINGS
# =========================
HOST = "localhost"
PORT = 10001
PASSWORD = "12345"

EXCEL_FILE = r"embedded_beams_top_results.xlsx"
PLOT_FILE_PNG = r"embedded_beams_scheme.png"

# =========================
# CONNECT TO OUTPUT
# =========================
s_o, g_o = new_server(HOST, PORT, password=PASSWORD)

# =========================
# HELPERS
# =========================
def get_last_phase(g_o):
    phase_names = [a for a in dir(g_o) if a.startswith("Phase_")]
    if phase_names:
        phase_names.sort(key=lambda x: int(x.split("_")[1]))
        return getattr(g_o, phase_names[-1])

    if hasattr(g_o, "Phases"):
        return g_o.Phases[-1]

    raise Exception("Cannot find phases in Output.")

def safe_name(obj, fallback):
    try:
        return obj.Name.value
    except:
        return fallback

def get_embedded_family(g_o):
    if hasattr(g_o.ResultTypes, "EmbeddedBeam"):
        return g_o.ResultTypes.EmbeddedBeam
    if hasattr(g_o.ResultTypes, "EmbeddedBeamRow"):
        return g_o.ResultTypes.EmbeddedBeamRow
    raise Exception("Embedded beam result type not found.")

def get_beams(g_o):
    if hasattr(g_o, "EmbeddedBeams"):
        return list(g_o.EmbeddedBeams)
    if hasattr(g_o, "EmbeddedBeamRows"):
        return list(g_o.EmbeddedBeamRows)
    raise Exception("No embedded beams found.")

def get_res(obj, phase, rtype):
    try:
        return g_o.getresults(obj, phase, rtype, "node")
    except:
        return []

def short_label(name, fallback_index):
    s = str(name)
    m = re.search(r'_(\d+)$', s)
    if m:
        return m.group(1)
    m = re.search(r'(\d+)$', s)
    if m:
        return m.group(1)
    return str(fallback_index)

def top_point_index(z_values):
    if not z_values:
        return None
    return max(range(len(z_values)), key=lambda i: z_values[i])

# =========================
# READ DATA
# =========================
phase = get_last_phase(g_o)
rt = get_embedded_family(g_o)
beams = get_beams(g_o)

beam_plot_data = []
top_results = []

for i, beam in enumerate(beams, start=1):
    original_name = safe_name(beam, f"Embedded beam_{i}")
    beam_id = short_label(original_name, i)

    X = get_res(beam, phase, rt.X)
    Y = get_res(beam, phase, rt.Y)
    Z = get_res(beam, phase, rt.Z)
    N = get_res(beam, phase, rt.N)
    Uz = get_res(beam, phase, rt.Uz) if hasattr(rt, "Uz") else []

    n_geom = min(len(X), len(Y), len(Z))
    if n_geom == 0:
        continue

    X = X[:n_geom]
    Y = Y[:n_geom]
    Z = Z[:n_geom]

    idx_top = top_point_index(Z)
    if idx_top is None:
        continue

    n_top = N[idx_top] if idx_top < len(N) else None
    uz_top = Uz[idx_top] if idx_top < len(Uz) else None

    beam_plot_data.append({
        "beam_id": beam_id,
        "original_name": original_name,
        "X": X,
        "Y": Y,
        "Z": Z,
        "top_x": X[idx_top],
        "top_y": Y[idx_top],
        "top_z": Z[idx_top],
        "N_top": n_top,
        "Uz_top": uz_top,
    })

    top_results.append([
        beam_id,
        original_name,
        X[idx_top],
        Y[idx_top],
        Z[idx_top],
        n_top,
        uz_top
    ])

# =========================
# EXPORT TO EXCEL
# =========================
wb = Workbook()

ws1 = wb.active
ws1.title = "Top_Results"
ws1.append(["BeamID", "OriginalName", "Top_X", "Top_Y", "Top_Z", "N_top", "Uz_top"])

for row in top_results:
    ws1.append(row)

for col in ["A", "B", "C", "D", "E", "F", "G"]:
    ws1.column_dimensions[col].width = 18

ws1.freeze_panes = "A2"

ws2 = wb.create_sheet("Geometry")
ws2.append(["BeamID", "OriginalName", "PointNo", "X", "Y", "Z"])

for beam in beam_plot_data:
    for j, (x, y, z) in enumerate(zip(beam["X"], beam["Y"], beam["Z"]), start=1):
        ws2.append([beam["beam_id"], beam["original_name"], j, x, y, z])

for col in ["A", "B", "C", "D", "E", "F"]:
    ws2.column_dimensions[col].width = 18

ws2.freeze_panes = "A2"

wb.save(EXCEL_FILE)

# =========================
# MATPLOTLIB GRAPHICAL OUTPUT
# =========================
plt.figure(figsize=(12, 9))

for beam in beam_plot_data:
    x = beam["X"]
    y = beam["Y"]
    top_x = beam["top_x"]
    top_y = beam["top_y"]
    beam_id = beam["beam_id"]

    # beam line
    plt.plot(x, y, linewidth=1.5)

    # top point
    plt.scatter([top_x], [top_y], s=35)

    # label beside top point
    plt.annotate(
        beam_id,
        (top_x, top_y),
        xytext=(5, 5),
        textcoords="offset points",
        fontsize=9
    )

plt.xlabel("X")
plt.ylabel("Y")
plt.title(f"Embedded beams scheme - top points labeled ({safe_name(phase, 'Last phase')})")
plt.grid(True)
plt.axis("equal")
plt.tight_layout()
plt.savefig(PLOT_FILE_PNG, dpi=300)
plt.show()

print("DONE")
print("Excel file:", EXCEL_FILE)
print("Plot file:", PLOT_FILE_PNG)
print("Beams exported:", len(top_results))
print("Phase:", safe_name(phase, "Last"))