using System.Diagnostics;

namespace FastDocView.Tests;

/// <summary>
/// Real OpenDocument test files: converts .docx files to .odt with a local LibreOffice (headless)
/// into TestOutput (gitignored — converted copies are never committed). Skipped when LibreOffice is
/// not installed; a conversion is reused while it is newer than its source.
/// </summary>
public static class LibreOfficeConversion
{
    /// <summary>LibreOffice asks Windows for the printer list and the default printer even headless; an offline
    /// default printer stalls the conversion. Both lookups are switched off by environment (vcl/source/gdi/print.cxx:372
    /// SAL_DISABLE_PRINTERLIST, :424 SAL_DISABLE_DEFAULTPRINTER). Every soffice start goes through here.</summary>
    internal static void WithoutPrinters(ProcessStartInfo start)
    {
        start.Environment["SAL_DISABLE_PRINTERLIST"] = "1";
        start.Environment["SAL_DISABLE_DEFAULTPRINTER"] = "1";
    }

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
                    Arguments = $"-env:UserInstallation=file:///{outputDirectory.Replace(Path.DirectorySeparatorChar, '/')}/lo-profile --headless --invisible --nologo --norestore --convert-to odt --outdir \"{outputDirectory}\" \"{Path.GetFullPath(docx)}\"",
                };
                LibreOfficeConversion.WithoutPrinters(start);
                using Process process = Process.Start(start)!;
                try { process.PriorityClass = ProcessPriorityClass.BelowNormal; } catch (InvalidOperationException) { }
                if (!process.WaitForExit(600_000)) { process.Kill(); TestReport.Check(false, "LibreOffice conversion finished within 10 minutes"); continue; }
            }
            TestReport.Check(File.Exists(odt), $"converted {Path.GetFileName(odt)} exists");
            if (File.Exists(odt)) converted.Add(odt);
        }
        return converted;
    }

    /// <summary>
    /// One file converted with a local LibreOffice (headless, private profile) to <paramref name="filter"/> (e.g.
    /// "doc:MS Word 97") into TestOutput; null when LibreOffice is not installed. Reused while newer than its source.
    /// </summary>
    public static string? Convert(string source, string extension, string filter)
    {
        if (!File.Exists(SofficePath)) return null;
        string outputDirectory = Path.Combine(TestReport.OutputDirectory, "lo-converted");
        Directory.CreateDirectory(outputDirectory);
        string target = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(source) + extension);
        if (File.Exists(target) && File.GetLastWriteTimeUtc(target) >= File.GetLastWriteTimeUtc(source)) return target;
        TestReport.Step($"Converting {Path.GetFileName(source)} to {extension} with LibreOffice (headless)");
        var start = new ProcessStartInfo(SofficePath)
        {
            UseShellExecute = false, CreateNoWindow = true,
            Arguments = $"-env:UserInstallation=file:///{outputDirectory.Replace(Path.DirectorySeparatorChar, '/')}/lo-profile --headless --invisible --nologo --norestore --convert-to \"{filter}\" --outdir \"{outputDirectory}\" \"{Path.GetFullPath(source)}\"",
        };
        LibreOfficeConversion.WithoutPrinters(start);
        using Process process = Process.Start(start)!;
        try { process.PriorityClass = ProcessPriorityClass.Idle; } catch (InvalidOperationException) { }
        if (!process.WaitForExit(300_000)) { process.Kill(); return null; }
        return File.Exists(target) ? target : null;
    }
}
