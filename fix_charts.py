"""
Fix the broken charts in the research instrument.
Issues:
1. Ray Structure scatter: noFill on line AND fill makes points invisible
2. Phase Transition scatter: same problem
3. Convergence line chart: no categories set
"""
import openpyxl
from openpyxl.chart import ScatterChart, Reference, Series, LineChart
from openpyxl.chart.marker import Marker
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.chart.shapes import GraphicalProperties
from copy import copy
import math

# Helper functions
def euler_totient(n):
    if n <= 0: return 0
    result = n; p = 2; temp = n
    while p * p <= temp:
        if temp % p == 0:
            while temp % p == 0: temp //= p
            result -= result // p
        p += 1
    if temp > 1: result -= result // temp
    return result

def divisor_count(n):
    if n <= 0: return 0
    count = 0
    for i in range(1, int(n**0.5) + 1):
        if n % i == 0:
            count += 1
            if i != n // i: count += 1
    return count

def d_func(n): return n - euler_totient(n) - divisor_count(n)

def is_prime(n):
    if n < 2: return False
    if n < 4: return True
    if n % 2 == 0 or n % 3 == 0: return False
    i = 5
    while i * i <= n:
        if n % i == 0 or n % (i + 2) == 0: return False
        i += 6
    return True

print("Loading workbook...")
wb = openpyxl.load_workbook(r'D:\claude\proofs\Deficiency_Research_Instrument.xlsx')

N_MAX = 500

# ═══════════════════════════════════════
# FIX SHEET 2: RAY STRUCTURE CHARTS
# ═══════════════════════════════════════
ws2 = wb['Ray Structure']

# Delete existing charts
ws2._charts = []

# CHART 1: d(n) vs n scatter — with VISIBLE markers
chart1 = ScatterChart()
chart1.title = "d(n) vs n — Ray Structure"
chart1.x_axis.title = "n"
chart1.y_axis.title = "d(n)"
chart1.style = 13
chart1.width = 28
chart1.height = 18

xvals = Reference(ws2, min_col=1, min_row=3, max_row=N_MAX + 2)
yvals = Reference(ws2, min_col=2, min_row=3, max_row=N_MAX + 2)
series1 = Series(yvals, xvals, title="d(n)")

# Make markers VISIBLE — small circles, blue
series1.marker = Marker(symbol='circle', size=3)
series1.graphicalProperties.line.noFill = True  # no connecting lines

chart1.series.append(series1)

# Y axis from -5 to max
chart1.y_axis.scaling.min = -5
chart1.y_axis.scaling.max = 350
chart1.x_axis.scaling.min = 0
chart1.x_axis.scaling.max = N_MAX + 10

ws2.add_chart(chart1, "F2")

# CHART 2: d(n)/n convergence as LINE chart with categories
chart2 = ScatterChart()
chart2.title = "d(n)/n Convergence to 1-6/pi^2 = 0.3921"
chart2.x_axis.title = "n"
chart2.y_axis.title = "d(n)/n"
chart2.style = 13
chart2.width = 28
chart2.height = 14

xvals2 = Reference(ws2, min_col=1, min_row=3, max_row=N_MAX + 2)
yvals2 = Reference(ws2, min_col=4, min_row=3, max_row=N_MAX + 2)
series2 = Series(yvals2, xvals2, title="d(n)/n")
series2.marker = Marker(symbol='circle', size=2)
series2.graphicalProperties.line.noFill = True

chart2.series.append(series2)

chart2.y_axis.scaling.min = -1.0
chart2.y_axis.scaling.max = 1.0
chart2.x_axis.scaling.min = 0
chart2.x_axis.scaling.max = N_MAX + 10

ws2.add_chart(chart2, "F25")

print("  Fixed Ray Structure charts")


# ═══════════════════════════════════════
# FIX SHEET 3: PHASE TRANSITION CHART
# ═══════════════════════════════════════
ws3 = wb['Phase Transition']

# Delete existing charts
ws3._charts = []

chart3 = ScatterChart()
chart3.title = "r(n) = (n - phi(n)) / tau(n) — Phase Transition at 1/2"
chart3.x_axis.title = "n"
chart3.y_axis.title = "r(n)"
chart3.style = 13
chart3.width = 28
chart3.height = 18

xvals3 = Reference(ws3, min_col=1, min_row=3, max_row=202)
yvals3 = Reference(ws3, min_col=4, min_row=3, max_row=202)
series3 = Series(yvals3, xvals3, title="r(n)")
series3.marker = Marker(symbol='circle', size=4)
series3.graphicalProperties.line.noFill = True  # dots only, no lines

chart3.series.append(series3)

chart3.y_axis.scaling.min = 0.0
chart3.y_axis.scaling.max = 6.0
chart3.x_axis.scaling.min = 0
chart3.x_axis.scaling.max = 205

ws3.add_chart(chart3, "G3")

print("  Fixed Phase Transition chart")


# ═══════════════════════════════════════
# ALSO FIX: FIBER HISTOGRAM
# ═══════════════════════════════════════
ws4 = wb['Fiber & Gap Analysis']
ws4._charts = []

from openpyxl.chart import BarChart
chart4 = BarChart()
chart4.title = "Fiber Sizes |F_k| for k = 0..100"
chart4.x_axis.title = "k"
chart4.y_axis.title = "|F_k|"
chart4.style = 13
chart4.width = 30
chart4.height = 14

# k=0 is row 4 (k=-1 is row 3), k=100 is row 104
cats4 = Reference(ws4, min_col=1, min_row=4, max_row=104)
vals4 = Reference(ws4, min_col=2, min_row=4, max_row=104)
chart4.add_data(vals4, titles_from_data=False)
chart4.set_categories(cats4)

ws4.add_chart(chart4, "I3")
print("  Fixed Fiber histogram")


# ═══════════════════════════════════════
# ALSO FIX: DIRICHLET CONVERGENCE
# ═══════════════════════════════════════
ws7 = wb['Dirichlet Series']
ws7._charts = []

chart7 = ScatterChart()
chart7.title = "Dirichlet Series |partial - formula| at s=3"
chart7.x_axis.title = "N (partial sum cutoff)"
chart7.y_axis.title = "|partial - formula|"
chart7.style = 13
chart7.width = 24
chart7.height = 14

xvals7 = Reference(ws7, min_col=1, min_row=3, max_row=12)
yvals7 = Reference(ws7, min_col=4, min_row=3, max_row=12)
s7 = Series(yvals7, xvals7, title="Error")
s7.marker = Marker(symbol='diamond', size=6)

chart7.series.append(s7)
chart7.x_axis.scaling.logBase = 10
chart7.y_axis.scaling.logBase = 10

ws7.add_chart(chart7, "A15")
print("  Fixed Dirichlet convergence chart")


# Save
out = r'D:\claude\proofs\Deficiency_Research_Instrument.xlsx'
wb.save(out)
print(f"\nSaved: {out}")
print("All 4 charts fixed. Reopen in Excel to see them.")
