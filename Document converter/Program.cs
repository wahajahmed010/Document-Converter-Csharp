using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Document_converter
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Global exception handlers
            Application.ThreadException += (sender, e) =>
            {
                Logger.Instance.Error($"Unhandled thread exception: {e.Exception.Message}", e.Exception);
                MessageBox.Show($"Unhandled error: {e.Exception.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                var ex = e.ExceptionObject as Exception;
                Logger.Instance.Error($"Fatal unhandled exception: {ex?.Message}", ex);
                if (e.IsTerminating)
                {
                    Environment.Exit(1);
                }
            };

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}