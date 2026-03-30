import sys
sys.path.append(r"C:\ProgramData\Seequent\PLAXIS Python Distribution V3\python\Lib\site-packages")

from plxscripting.easy import *
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Series, Reference

# =========================
# SETTINGS
# =========================
HOST = "localhost"
PORT = 10001   # ✅ OUTPUT PORT (fixed)
PASSWORD = "12345"
OUT_FILE = r"embedded_beams_last_phase.xlsx"

# =========================
# CONNECT TO OUTPUT
# =========================
s_o, g_o = new_server(HOST, PORT, password=PASSWORD)

# =========================
# GET LAST PHASE (robust)
# =========================
def get_last_phase(g_o):
    # Newer versions (Output)
    phase_names = [a for a in dir(g_o) if a.startswith("Phase_")]
    if phase_names:
        phase_names.sort(key=lambda x: int(x.split("_")[1]))
        return getattr(g_o, phase_names[-1])

    # Older versions fallback
    if hasattr(g_o, "Phases"):
        return g_o.Phases[-1]

    raise Exception("Cannot find phases in Output")

phase = get_last_phase(g_o)

# =========================
# HELPERS
# =========================
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
    raise Exception("Embedded beam result type not found")

def get_beams(g_o):
    if hasattr(g_o, "EmbeddedBeams"):
        return list(g_o.EmbeddedBeams)
    if hasattr(g_o, "EmbeddedBeamRows"):
        return list(g_o.EmbeddedBeamRows)
    raise Exception("No embedded beams found")

def get_res(obj, phase, rtype):
    try:
        return g_o.getresults(obj, phase, rtype, "node")
    except:
        return []

# =========================
# EXTRACTION
# =========================
rt = get_embedded_family(g_o)
beams = get_beams(g_o)

wb = Workbook()
ws = wb.active
ws.title = "All_Beams"

ws.append(["Beam", "Point", "X", "Y", "Z", "N", "Uz"])

exported = 0

for i, beam in enumerate(beams, start=1):
    name = safe_name(beam, f"Beam_{i}")

    X = get_res(beam, phase, rt.X)
    Y = get_res(beam, phase, rt.Y)
    Z = get_res(beam, phase, rt.Z)
    N = get_res(beam, phase, rt.N)

    # Uz can vary by version → safe handling
    Uz = get_res(beam, phase, rt.Uz) if hasattr(rt, "Uz") else []

    npts = min(len(X), len(Y), len(Z), len(N), len(Uz) if Uz else len(X))
    if npts == 0:
        continue

    exported += 1

    for j in range(npts):
        uz_val = Uz[j] if Uz else None
        ws.append([name, j+1, X[j], Y[j], Z[j], N[j], uz_val])

# =========================
# FORMAT
# =========================
for col in "ABCDEFG":
    ws.column_dimensions[col].width = 16

ws.freeze_panes = "A2"

# =========================
# XY SCHEME
# =========================
chart = ScatterChart()
chart.title = f"Embedded beams (Last phase: {safe_name(phase,'')})"
chart.x_axis.title = "X"
chart.y_axis.title = "Y"
chart.scatterStyle = "lineMarker"
chart.height = 14
chart.width = 24

max_row = ws.max_row
r = 2

while r <= max_row:
    beam_name = ws.cell(r, 1).value
    r_start = r

    while r <= max_row and ws.cell(r, 1).value == beam_name:
        r += 1

    r_end = r - 1

    x_ref = Reference(ws, min_col=3, min_row=r_start, max_row=r_end)
    y_ref = Reference(ws, min_col=4, min_row=r_start, max_row=r_end)

    chart.series.append(Series(y_ref, x_ref, title=beam_name))

ws.add_chart(chart, "I2")

# =========================
# SAVE
# =========================
wb.save(OUT_FILE)

print("DONE")
print("File:", OUT_FILE)
print("Beams exported:", exported)
print("Phase:", safe_name(phase, "Last"))