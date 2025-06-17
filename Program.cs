using System;
using System.IO;
using System.Linq;

namespace SilentSlideUpdater
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
               
               DateTime today = DateTime.Today;
               DateTime firstTuesday = new DateTime(today.Year, today.Month, 1);
               while (firstTuesday.DayOfWeek != DayOfWeek.Tuesday)
                    firstTuesday = firstTuesday.AddDays(1);
                
                if (today != firstTuesday)
                    return;

                // Set up paths
                string basePath = @"\\<INTERNAL_IP>\information system\Campaign\FY25\Security Poster";
                string currentSlide = @"\\<INTERNAL_IP>\Campaign\SecurityPoster.ppsx";

                if (!File.Exists(currentSlide))

                    return;

                byte[] currentContent = File.ReadAllBytes(currentSlide);

                var numberedDirs = Directory.GetDirectories(basePath)
                    .Select(Path.GetFileName)
                    .Where(name => int.TryParse(name, out _))
                    .Select(int.Parse)
                    .OrderBy(n => n)
                    .ToList();

                int matchIndex = -1;

                for (int i = 0; i < numberedDirs.Count; i++)
                {
                    string folder = Path.Combine(basePath, numberedDirs[i].ToString());
                    var ppsxFile = Directory.GetFiles(folder, "*.ppsx").FirstOrDefault();
                    if (ppsxFile == null) continue;

                    byte[] compareContent = File.ReadAllBytes(ppsxFile);
                    if (currentContent.SequenceEqual(compareContent))
                    {
                        matchIndex = i;
                        break;
                    }
                }

                if (matchIndex == -1 || matchIndex + 1 >= numberedDirs.Count)
                    return;

                int nextFolder = numberedDirs[matchIndex + 1];
                string nextPpsx = Directory.GetFiles(Path.Combine(basePath, nextFolder.ToString()), "*.ppsx").FirstOrDefault();

                if (nextPpsx == null)
                    return;

                File.Copy(nextPpsx, currentSlide, overwrite: true);
            }
            catch (Exception ex)
            {
                File.AppendAllText(@"C:\Temp\SlideUpdateLog.txt", $"[{DateTime.Now}] ERROR: {ex.Message}\n");
            }
        }
    }
}
