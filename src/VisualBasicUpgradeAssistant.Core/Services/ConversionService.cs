using System;
using System.Diagnostics;
using System.IO;

namespace VisualBasicUpgradeAssistant.Core.Services
{
    internal class ConversionService
    {
        public void Convert(FileInfo source, DirectoryInfo dest)
        {
            String basename = Path.GetFileNameWithoutExtension(source.Name);

            CreateProject(dest, basename);

            // fill the files.
        }

        private static void CreateProject(DirectoryInfo dest, String basename)
        {
            // Setup output Dir
            DirectoryInfo outputDir = new DirectoryInfo(Path.Combine(dest.FullName, basename));
            if (outputDir.Exists)
                outputDir.Delete(true);

            // setup for initial project
            ProcessStartInfo startinfo = new ProcessStartInfo
            {
                CreateNoWindow = true,
                UseShellExecute = false,
                WorkingDirectory = dest.FullName,
                FileName = "dotnet.exe",
                Arguments = $"new sln -o {basename}",
            };

            // setup base project
            Process.Start(startinfo).WaitForExit();

            // Add library
            startinfo.WorkingDirectory = outputDir.FullName;
            startinfo.Arguments = $"new classlib -o {basename}.Core";
            Process.Start(startinfo).WaitForExit();

            // Add Test
            startinfo.Arguments = $"new mstest -o {basename}.CoreTests";
            Process.Start(startinfo).WaitForExit();

            // Add Test
            startinfo.Arguments = $"add {basename}.CoreTests/{basename}.CoreTests.csproj reference {basename}.Core/{basename}.Core.csproj";
            Process.Start(startinfo).WaitForExit();

            // Add Winforms
            startinfo.Arguments = $"new winforms -o {basename}.Desktop";
            Process.Start(startinfo).WaitForExit();

            // Add Test
            startinfo.Arguments = $"add {basename}.Desktop/{basename}.Desktop.csproj reference {basename}.Core/{basename}.Core.csproj";
            Process.Start(startinfo).WaitForExit();

            // Combine into sln
            startinfo.Arguments = $"sln {basename}.sln add {basename}.Core {basename}.CoreTests {basename}.Desktop";
            Process.Start(startinfo).WaitForExit();
        }
    }
}
