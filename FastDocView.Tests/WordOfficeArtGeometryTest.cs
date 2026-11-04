using FastDocView.Html;

namespace FastDocView.Tests;

/// <summary>Every OfficeArt preset LibreOffice has a table for evaluates to finite outlines inside (or near) its box.</summary>
public static class WordOfficeArtGeometryTest
{
    public static void Run()
    {
        TestReport.Step("OfficeArt presets: every table evaluates");
        int drawn = 0, bad = 0;
        var problems = new List<int>();
        for (int type = 0; type < 256; type++)
        {
            List<WordShapePath>? paths = WordOfficeArtGeometry.Preset(type, 100, 50, []);
            if (paths is null) continue;
            drawn++;
            bool ok = paths.Count > 0;
            foreach (WordShapePath path in paths)
                foreach (WordShapeSubpath subpath in path.Subpaths)
                {
                    ok &= double.IsFinite(subpath.StartX) && double.IsFinite(subpath.StartY);
                    foreach (WordShapeSegment segment in subpath.Segments)
                        ok &= double.IsFinite(segment.X) && double.IsFinite(segment.Y) && Math.Abs(segment.X) < 1000 && Math.Abs(segment.Y) < 1000;
                }
            if (!ok) { bad++; problems.Add(type); }
        }
        TestReport.Check(drawn >= 180, $"{drawn} preset types have outlines");
        TestReport.Check(bad == 0, $"every outline is finite and near its box (problems: {string.Join(", ", problems)})");
    }
}
