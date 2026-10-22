using Cleaning_Records_Sheet;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Forms;

public static class ChromeDriverManager
{
    private static string networkFolder;
    private static string appFolder;

    static ChromeDriverManager()
    {
        appFolder = AppDomain.CurrentDomain.BaseDirectory;

        // Read network path from text file
        string configPath = Path.Combine(appFolder, "bin", "networkfolder.txt");

        if (File.Exists(configPath))
        {
            networkFolder = File.ReadAllText(configPath);
        }
        else
        {
            throw new Exception("networkfolder.txt not found.");
        }
    }
   
    public static bool CheckDriver()
    {
        try
        {
            if (Utility.m_eRUN == RUN.CYIENT) return true;
            // Get Chrome Browser Path
            string chromePath = @"C:\Program Files\Google\Chrome\Application\chrome.exe";
            if (!File.Exists(chromePath))
            {
                MessageBox.Show("Google Chrome is not installed.");
                return false;
            }

            // Get Chrome Browser Version
            string chromeVersion = FileVersionInfo.GetVersionInfo(chromePath).FileVersion;
            string chromeMajor = chromeVersion.Split('.')[0];

            //Get Chrome Driver Path
            string chromeDriverPath = Path.Combine(appFolder, "chromedriver.exe");

            //Get Chrome Driver Version
            string driverVersion = FileVersionInfo.GetVersionInfo(chromeDriverPath).FileVersion;
            string driverMajor = driverVersion.Split('.')[0];

            if (File.Exists(chromeDriverPath))
            {
                //Check if chrome browser version and chrome driver version match
                if (chromeMajor == driverMajor)
                {
                    return true;
                }
            }

            // Check in Network Folder if no match found
            Version chromeVer = new Version(chromeVersion);
            string selectedFolder = null;
            Version nearestVersion = null;

            foreach (var dir in Directory.GetDirectories(networkFolder))
            {
                string folderName = Path.GetFileName(dir);
                string versionPart = folderName.Split(new char[] { '~', '_', '-' })[0];

                if (Version.TryParse(versionPart, out Version folderVersion))
                {
                    if (folderVersion <= chromeVer)
                    {
                        if (nearestVersion == null || folderVersion > nearestVersion)
                        {
                            nearestVersion = folderVersion;
                            selectedFolder = dir;
                        }
                    }
                }
            }

            // If not found in network folder give popup message
            if (selectedFolder == null)
            {
                MessageBox.Show(
                 "ChromeDriverのアップデートが必要です。DX担当者にお知らせください。\n" +
                 $"ChromeDriver version: {driverVersion}\n" +
                 $"Chrome browser version: {chromeVersion}"
                );

                return false;
            }

            //If found copy from network folder to app folder
            string sourceDriver = Path.Combine(selectedFolder, "chromedriver.exe");
            string destinationDriver = Path.Combine(appFolder, "chromedriver.exe");

            if (!File.Exists(sourceDriver))
            {
                MessageBox.Show("chromedriver.exe not found in Network Folder.");
                return false;
            }

            File.Copy(sourceDriver, destinationDriver, true);

            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(ex.Message);
            return false;
        }
    }
}