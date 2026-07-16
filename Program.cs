namespace DoseXPDMTool
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += (_, args) => LogApplicationException("UI thread exception", args.Exception);
            AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            {
                if (args.ExceptionObject is Exception ex)
                {
                    LogApplicationException("Unhandled app exception", ex);
                }
            };
            TaskScheduler.UnobservedTaskException += (_, args) =>
            {
                LogApplicationException("Unobserved task exception", args.Exception);
                args.SetObserved();
            };
            Application.Run(new DoseX_Point_Dose_Tool());
        }

        private static void LogApplicationException(string context, Exception ex)
        {
            try
            {
                string directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "DoseXPDMTool");
                Directory.CreateDirectory(directory);
                string logPath = Path.Combine(directory, "DoseXPDMTool-errors.log");
                string message = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}{Environment.NewLine}{ex}{Environment.NewLine}";
                File.AppendAllText(logPath, message);
            }
            catch
            {
            }
        }
    }
}
