using System.Diagnostics;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Real OpenDocument test files: converts .docx files to .odt with a local LibreOffice (headless)
/// into TestOutput (gitignored — converted copies are never committed). Skipped when LibreOffice is
/// not installed; a conversion is reused while it is newer than its source.
/// </summary>
public static class LibreOfficeConversion
{
    private const string SofficePath = @"C:\Program Files\LibreOffice\program\soffice.exe";

    public static List<string> ConvertToOdt(IEnumerable<string> docxPaths)
    {
        var converted = new List<string>();
        if (!File.Exists(SofficePath)) { TestReport.Step("LibreOffice not installed — no real .odt conversion"); return converted; }
        string outputDirectory = Path.Combine(TestReport.OutputDirectory, "odt-converted");
        Directory.CreateDirectory(outputDirectory);
        foreach (string docx in docxPaths)
        {
            string odt = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(docx) + ".odt");
            if (!File.Exists(odt) || File.GetLastWriteTimeUtc(odt) < File.GetLastWriteTimeUtc(docx))
            {
                TestReport.Step($"Converting {Path.GetFileName(docx)} to .odt with LibreOffice (headless)");
                var start = new ProcessStartInfo(SofficePath)
                {
                    UseShellExecute = false, CreateNoWindow = true,
                    // a private profile so a running LibreOffice of the user is neither used nor disturbed
                    Arguments = $"-env:UserInstallation=file:///{outputDirectory.Replace(Path.DirectorySeparatorChar, '/')}/lo-profile --headless --norestore --convert-to odt --outdir \"{outputDirectory}\" \"{Path.GetFullPath(docx)}\"",
                };
                using Process process = Process.Start(start)!;
                try { process.PriorityClass = ProcessPriorityClass.BelowNormal; } catch (InvalidOperationException) { }
                if (!process.WaitForExit(600_000)) { process.Kill(); TestReport.Check(false, "LibreOffice conversion finished within 10 minutes"); continue; }
            }
            TestReport.Check(File.Exists(odt), $"converted {Path.GetFileName(odt)} exists");
            if (File.Exists(odt)) converted.Add(odt);
        }
        return converted;
    }
}
