from plxscripting.easy import *
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Series, Reference

# =========================
# SETTINGS
# =========================
HOST = "localhost"
PORT = 10000          # Output port
PASSWORD = "12345"
OUT_FILE = r"embedded_beams_last_phase.xlsx"

# =========================
# CONNECT TO PLAXIS OUTPUT
# =========================
s_o, g_o = new_server(HOST, PORT, password=PASSWORD)

# Last phase only
phase = g_o.Phases[-1]

# =========================
# HELPERS
# =========================
def get_embedded_result_family(g_o):
    if hasattr(g_o.ResultTypes, "EmbeddedBeam"):
        return g_o.ResultTypes.EmbeddedBeam
    if hasattr(g_o.ResultTypes, "EmbeddedBeamRow"):
        return g_o.ResultTypes.EmbeddedBeamRow
    raise Exception("Could not find EmbeddedBeam / EmbeddedBeamRow in ResultTypes.")

def safe_name(obj, fallback):
    try:
        return obj.Name.value
    except:
        return fallback

def get_all_embedded_beams(g_o):
    if hasattr(g_o, "EmbeddedBeams"):
        return list(g_o.EmbeddedBeams)
    if hasattr(g_o, "EmbeddedBeamRows"):
        return list(g_o.EmbeddedBeamRows)
    raise Exception("Could not find EmbeddedBeams / EmbeddedBeamRows collection.")

def get_result(g_o, obj, phase, result_type, result_location="node"):
    try:
        return g_o.getresults(obj, phase, result_type, result_location)
    except:
        return []

# =========================
# READ RESULTS
# =========================
rt = get_embedded_result_family(g_o)
beams = get_all_embedded_beams(g_o)

wb = Workbook()
ws = wb.active
ws.title = "All_Beams"

headers = [
    "BeamName",
    "PointNo",
    "X",
    "Y",
    "Z",
    "N",
    "Uz",
]
ws.append(headers)

current_row = 2
exported_beams = 0

for i, beam in enumerate(beams, start=1):
    beam_name = safe_name(beam, f"Beam_{i}")

    x_vals = get_result(g_o, beam, phase, rt.X, "node")
    y_vals = get_result(g_o, beam, phase, rt.Y, "node")
    z_vals = get_result(g_o, beam, phase, rt.Z, "node")
    n_vals = get_result(g_o, beam, phase, rt.N, "node")
    uz_vals = get_result(g_o, beam, phase, rt.Uz, "node")

    npts = min(len(x_vals), len(y_vals), len(z_vals), len(n_vals), len(uz_vals))
    if npts == 0:
        continue

    exported_beams += 1

    for j in range(npts):
        ws.append([
            beam_name,
            j + 1,
            x_vals[j],
            y_vals[j],
            z_vals[j],
            n_vals[j],
            uz_vals[j],
        ])
        current_row += 1

# =========================
# FORMAT
# =========================
for col in ["A", "B", "C", "D", "E", "F", "G"]:
    ws.column_dimensions[col].width = 18

ws.freeze_panes = "A2"

# =========================
# CREATE XY SCHEME CHART
# =========================
chart = ScatterChart()
chart.title = f"Embedded beams scheme - {safe_name(phase, 'Last phase')}"
chart.style = 2
chart.x_axis.title = "X"
chart.y_axis.title = "Y"
chart.height = 14
chart.width = 24
chart.scatterStyle = "lineMarker"

last_data_row = ws.max_row
r = 2
while r <= last_data_row:
    beam_name = ws.cell(r, 1).value
    r_start = r
    while r <= last_data_row and ws.cell(r, 1).value == beam_name:
        r += 1
    r_end = r - 1

    xvalues = Reference(ws, min_col=3, min_row=r_start, max_row=r_end)  # X
    yvalues = Reference(ws, min_col=4, min_row=r_start, max_row=r_end)  # Y
    series = Series(yvalues, xvalues, title=beam_name)
    chart.series.append(series)

ws.add_chart(chart, "I2")

# =========================
# SAVE
# =========================
wb.save(OUT_FILE)

print(f"Done. File saved: {OUT_FILE}")
print(f"Phase used: {safe_name(phase, 'Last phase')}")
print(f"Number of embedded beams exported: {exported_beams}")