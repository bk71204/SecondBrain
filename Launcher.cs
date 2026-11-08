using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

internal static class Launcher
{
    [STAThread]
    private static void Main()
    {
        string scriptPath = null;
        try
        {
            string script;
            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream("ExportarAObsidian.ps1"))
            {
                if (stream == null)
                    throw new InvalidOperationException("No se encontró el script de la aplicación.");
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true))
                    script = reader.ReadToEnd();
            }

            scriptPath = Path.Combine(
                Path.GetTempPath(),
                "ExportarAObsidian-" + Guid.NewGuid().ToString("N") + ".ps1");
            File.WriteAllText(scriptPath, script, new UTF8Encoding(false));
            string powershell = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.System),
                @"WindowsPowerShell\v1.0\powershell.exe");
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = powershell,
                Arguments = "-NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File \"" + scriptPath + "\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            using (Process process = Process.Start(startInfo))
                process.WaitForExit();
        }
        catch (Exception exception)
        {
            MessageBox.Show(exception.Message, "Exportar a Obsidian", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (scriptPath != null)
            {
                try
                {
                    File.Delete(scriptPath);
                }
                catch (Exception exception)
                {
                    MessageBox.Show(
                        "No se pudo eliminar el script temporal: " + exception.Message,
                        "Exportar a Obsidian",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
        }
    }
}
