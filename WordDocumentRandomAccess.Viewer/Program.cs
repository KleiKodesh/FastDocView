namespace WordDocumentRandomAccess.Viewer;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        string? filePath = null, screenshotPath = null;
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--screenshot" && i + 1 < args.Length) screenshotPath = args[++i];
            else if (!args[i].StartsWith("--")) filePath = args[i];
        }
        ApplicationConfiguration.Initialize();
        Application.Run(new ViewerMainForm(filePath, screenshotPath));
    }
}
