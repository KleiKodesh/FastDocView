# Makes two fixtures from word-charts-binary-v1.docx with the helpers of make-word-charts-seraxis-v1.py (chart XML rewritten
# directly, no Excel automation):
#   word-charts-right-angled-v1.docx - right-angled 3-D areas / columns at several rotX / rotY, with and without a shown series
#                                      axis: where Word puts the scene and the series names
#   word-charts-theme-text-v1.docx   - chart text in theme colours (accents with lumMod / lumOff) under a document theme whose
#                                      accents are not Office's defaults: the .doc / .rtf copies keep them exactly only in their
#                                      DrawingML records plus the theme Office zips into a CrtMlFrt record
# Then scripts/save-word-charts-seraxis.ps1 lets Word save each as PDF, .doc and .rtf.
from importlib import util
import os
spec = util.spec_from_file_location('seraxis', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'make-word-charts-seraxis-v1.py'))
seraxis = util.module_from_spec(spec); spec.loader.exec_module(seraxis)
make, write = seraxis.make, seraxis.write

def right(rot_x, rot_y):
    return {'rotX': rot_x, 'rotY': rot_y, 'rAngAx': 1}

write('word-charts-right-angled-v1.docx', {
    1: make('area3d', 'R 20 30 S', right(20, 30)),
    2: make('area3d', 'R 20 30', right(20, 30), series_axis=False),
    3: make('area3d', 'R 30 20', right(30, 20), series_axis=False),
    4: make('area3d', 'R 15 40', right(15, 40), series_axis=False),
    5: make('area3d', 'R 5 30', right(5, 30), series_axis=False),
    6: make('area3d', 'R 15 20', right(15, 20), series_axis=False),
    7: make('col3d', 'C 25 35 S', right(25, 35)),
})

def scheme(name, mod=None, off=None):
    inner = ('<a:lumMod val="%d"/>' % mod if mod else '') + ('<a:lumOff val="%d"/>' % off if off else '')
    return '<a:schemeClr val="%s">%s</a:schemeClr>' % (name, inner)

write('word-charts-theme-text-v1.docx', {
    1: make('col', 'Theme Text', None, label_color=scheme('accent2', 75000), title_color=scheme('accent6'), legend_color=scheme('accent1', 60000, 40000)),
    2: make('col3d', 'Theme Text 3D', seraxis.RIGHT, label_color=scheme('accent5', 50000), title_color=scheme('tx2'), legend_color=scheme('accent4')),
}, accents={'accent1': '1F8A70', 'accent2': '9C27B0', 'accent4': 'C2185B', 'accent5': '3949AB', 'accent6': 'E65100', 'dk2': '5D4037'})
