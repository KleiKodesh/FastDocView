# Builds test-data/odt-charts/odt-charts-v4.odt (and, converted by LibreOffice, odt-charts-v4.docx) with LibreOffice itself (UNO): one Writer chart per chart type.
# Run with LibreOffice's python while soffice listens:  soffice --headless --accept="pipe,name=fdvcharts;urp;" (SAL_DISABLE_PRINTERLIST=1 SAL_DISABLE_DEFAULTPRINTER=1)
import sys, time, uno
from com.sun.star.beans import PropertyValue

out = sys.argv[1]
local = uno.getComponentContext()
resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
for _ in range(60):
    try:
        ctx = resolver.resolve("uno:pipe,name=fdvcharts;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(1)
desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
hidden = PropertyValue(); hidden.Name = "Hidden"; hidden.Value = True
doc = desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, (hidden,))
text = doc.Text
cursor = text.createTextCursor()

def extras(name, chart):
    d = chart.Diagram
    if name == "ChartLog":
        d.getYAxis().Logarithmic = True
    if name == "ChartSecondary":
        d.HasSecondaryYAxis = True
        d.getDataRowProperties(1).Axis = 4  # ChartAxisAssign.SECONDARY_Y
        d.HasSecondaryYAxisTitle = True; d.SecondYAxisTitle.String = "Ratio Axis"
    if name in ("ChartErrors", "ChartRangeErrors"):
        row = d.getDataRowProperties(0)
        row.ErrorCategory = uno.Enum("com.sun.star.chart.ChartErrorCategory", "CONSTANT_VALUE")
        row.ConstantErrorHigh = 2.0; row.ConstantErrorLow = 1.0
        row.ErrorIndicator = uno.Enum("com.sun.star.chart.ChartErrorIndicatorType", "TOP_AND_BOTTOM")
    if name == "ChartTrend":
        series = chart.getFirstDiagram().getCoordinateSystems()[0].getChartTypes()[0].getDataSeries()[0]
        curve = ctx.ServiceManager.createInstanceWithContext("com.sun.star.chart2.LinearRegressionCurve", ctx)
        series.addRegressionCurve(curve)
        added = series.getRegressionCurves()[0]
        for key, value in (("ExtrapolateForward", 1.0), ("ExtrapolateBackward", 0.5), ("ForceIntercept", True), ("InterceptValue", 0.0)):
            try: added.setPropertyValue(key, value)
            except Exception as e: print("curve", key, e)
    if name == "ChartWrapped":
        d.getXAxis().TextBreak = True
    if name == "ChartColumn3D":
        d.getFloor().FillColor = 0xCCE5FF

CHARTS = [
    # (name, diagram service, props, title, rows, columns, data)
    ("ChartColumn", "com.sun.star.chart.BarDiagram", {}, "Column Sales", ["Q1", "Q2", "Q3"], ["Alpha", "Beta"], [[10, 20], [30, 15], [40, 25]]),
    ("ChartBarStacked", "com.sun.star.chart.BarDiagram", {"Vertical": True, "Stacked": True}, "Stacked Bars", ["North", "South"], ["Apples", "Pears"], [[3, 4], [5, 1]]),
    ("ChartPercent", "com.sun.star.chart.BarDiagram", {"Percent": True}, "Percent Columns", ["One", "Two"], ["Left", "Right"], [[1, 3], [2, 2]]),
    ("ChartLine", "com.sun.star.chart.LineDiagram", {}, "Line Trend", ["Jan", "Feb", "Mar"], ["Trend"], [[1], [3], [2]]),
    ("ChartArea", "com.sun.star.chart.AreaDiagram", {"Stacked": True}, "Area Stack", ["A", "B", "C"], ["Lower", "Upper"], [[2, 2], [2, 6], [3, 1]]),
    ("ChartPie", "com.sun.star.chart.PieDiagram", {}, "Pie Share", ["East", "West", "Middle"], ["Share"], [[50], [25], [25]]),
    ("ChartDonut", "com.sun.star.chart.DonutDiagram", {}, "Donut Rings", ["Red", "Blue"], ["Ring"], [[60], [40]]),
    ("ChartScatter", "com.sun.star.chart.XYDiagram", {}, "Scatter Points", ["1", "2", "3"], ["X", "Points"], [[1, 2], [2, 4], [3, 3]]),
    ("ChartRadar", "com.sun.star.chart.NetDiagram", {}, "Radar Net", ["Speed", "Power", "Range", "Cost"], ["Car", "Bike"], [[4, 2], [3, 1], [5, 3], [2, 4]]),
    ("ChartStock", "com.sun.star.chart.StockDiagram", {}, "Stock Prices", ["Mon", "Tue", "Wed"], ["Low", "High", "Close"], [[10, 14, 12], [11, 15, 13], [9, 13, 10]]),
    ("ChartBubble", "com.sun.star.chart.BubbleDiagram", {}, "Bubble Sizes", ["1", "2"], ["X", "Y", "Size"], [[1, 2, 3], [2, 3, 5]]),
    ("ChartColumn3D", "com.sun.star.chart.BarDiagram", {"Dim3D": True}, "Column Depth", ["P", "Q"], ["Solid"], [[4], [7]]),
    ("ChartLog", "com.sun.star.chart.LineDiagram", {}, "Log Growth", ["A", "B", "C", "D"], ["Growth"], [[1], [10], [100], [1000]]),
    ("ChartSecondary", "com.sun.star.chart.BarDiagram", {}, "Two Axes", ["J", "K", "L"], ["Units", "Ratio"], [[100, 0.2], [200, 0.5], [300, 0.9]]),
    ("ChartErrors", "com.sun.star.chart.BarDiagram", {}, "Error Bars", ["E1", "E2", "E3"], ["Measured"], [[5], [8], [6]]),
    ("ChartTrend", "com.sun.star.chart.XYDiagram", {}, "Trend Fit", ["1", "2", "3", "4"], ["X", "Obs"], [[1, 2], [2, 4.5], [3, 5.5], [4, 8]]),
    ("ChartPie3D", "com.sun.star.chart.PieDiagram", {"Dim3D": True}, "Pie Depth", ["North", "South", "East", "West"], ["Share"], [[40], [30], [20], [10]]),
    ("ChartRangeErrors", "com.sun.star.chart.BarDiagram", {}, "Range Errors", ["R1", "R2", "R3"], ["Level"], [[2], [3], [4]]),
    ("ChartLine3D", "com.sun.star.chart.LineDiagram", {"Dim3D": True}, "Line Depth", ["A", "B", "C", "D"], ["Front", "Back"], [[2, 4], [5, 3], [3, 6], [6, 5]]),
    ("ChartArea3D", "com.sun.star.chart.AreaDiagram", {"Dim3D": True}, "Area Depth", ["A", "B", "C"], ["Low", "High"], [[2, 5], [4, 7], [3, 6]]),
    ("ChartBar3DH", "com.sun.star.chart.BarDiagram", {"Dim3D": True, "Vertical": True}, "Bar Depth", ["Left", "Right"], ["Size"], [[3], [6]]),
    ("ChartCrowded", "com.sun.star.chart.BarDiagram", {}, "Crowded Axis", ["Categorylabel%02d" % i for i in range(1, 13)], ["Count"], [[i % 5 + 1] for i in range(12)]),
    ("ChartWrapped", "com.sun.star.chart.BarDiagram", {}, "Wrapped Labels", ["North east region", "South west region", "Central city area", "Outer ring towns", "Island and coast", "Mountain villages"], ["Count"], [[3], [5], [4], [6], [2], [5]]),
    ("ChartCandle", "com.sun.star.chart.StockDiagram", {"UpDown": True}, "Candle Stock", ["D1", "D2", "D3"], ["Open", "Low", "High", "Close"], [[10, 8, 14, 13], [13, 11, 15, 12], [12, 9, 16, 15]]),
]
for name, service, props, title, rows, cols, data in CHARTS:
    text.insertString(cursor, name + " paragraph", False)
    obj = doc.createInstance("com.sun.star.text.TextEmbeddedObject")
    obj.CLSID = "12dcae26-281f-416f-a234-c3086127382e"
    obj.Name = name
    size = uno.createUnoStruct("com.sun.star.awt.Size"); size.Width = 12000; size.Height = 7000
    obj.Size = size
    obj.AnchorType = uno.Enum("com.sun.star.text.TextContentAnchorType", "AS_CHARACTER")
    text.insertTextContent(cursor, obj, False)
    chart = obj.getEmbeddedObject()
    chart.lockControllers()
    diagram = chart.createInstance(service)
    chart.setDiagram(diagram)
    data_array = chart.getData()
    data_array.setData(tuple(tuple(float(v) for v in row) for row in data))
    data_array.setRowDescriptions(tuple(rows))
    data_array.setColumnDescriptions(tuple(cols))
    for key, value in props.items():
        try: setattr(chart.Diagram, key, value)
        except Exception as e: print("prop", name, key, e)
    chart.HasMainTitle = True
    chart.Title.String = title
    if name == "ChartColumn":
        chart.HasSubTitle = True; chart.SubTitle.String = "Sub Heading"
        chart.Diagram.HasXAxisTitle = True; chart.Diagram.XAxisTitle.String = "Quarter Axis"
        chart.Diagram.HasYAxisTitle = True; chart.Diagram.YAxisTitle.String = "Amount Axis"
        chart.Diagram.DataCaption = 1  # VALUE
    extras(name, chart)
    chart.HasLegend = True
    chart.unlockControllers()
    text.insertControlCharacter(cursor, 0, False)  # PARAGRAPH_BREAK

url = uno.systemPathToFileUrl(out)
f = PropertyValue(); f.Name = "FilterName"; f.Value = "writer8"
doc.storeToURL(url, (f,))
doc.close(True)
# the trend line's equation and R squared: written into the chart's XML (the equation properties are not reachable through UNO
# from Python here); LibreOffice reads chart:equation back (xmloff SchXMLRegressionCurveObjectContext.cxx:81-136)
import zipfile, shutil, os
temporary = out + ".tmp"
with zipfile.ZipFile(out) as source, zipfile.ZipFile(temporary, "w") as target:
    for item in source.infolist():
        data = source.read(item.filename)
        if item.filename.endswith("/content.xml") and b"Range Errors" in data:
            # error amounts from a cell range (ODF chart:error-category cell-range, error-upper-range / lower-range): the values column itself
            text = data.decode("utf-8").replace('chart:error-category="constant"', 'chart:error-category="cell-range" chart:error-upper-range="local-table.$B$2:.$B$4" chart:error-lower-range="local-table.$B$2:.$B$4"')
            data = text.encode("utf-8")
        if item.filename.endswith("/content.xml") and b"chart:regression-curve" in data:
            text = data.decode("utf-8")
            start = text.index("<chart:regression-curve")
            close = text.index(">", start)
            if text[close - 1] == "/":
                text = text[:close - 1] + '><chart:equation chart:display-equation="true" chart:display-r-square="true"/></chart:regression-curve>' + text[close + 1:]
            data = text.encode("utf-8")
        target.writestr(item, data, compress_type=zipfile.ZIP_STORED if item.filename == "mimetype" else zipfile.ZIP_DEFLATED)
shutil.move(temporary, out)
print("saved", out)
