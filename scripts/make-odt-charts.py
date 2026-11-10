# Builds test-data/odt-charts/odt-charts-v1.odt with LibreOffice itself (UNO): one Writer chart per chart type.
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
    chart.HasLegend = True
    chart.unlockControllers()
    text.insertControlCharacter(cursor, 0, False)  # PARAGRAPH_BREAK

url = uno.systemPathToFileUrl(out)
f = PropertyValue(); f.Name = "FilterName"; f.Value = "writer8"
doc.storeToURL(url, (f,))
doc.close(True)
print("saved", out)
