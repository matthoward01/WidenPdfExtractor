using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Xml;

namespace WidenPdfExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            Settings settings = GetSettings();
            //string inFile = @"/Users/matt.howard/Desktop/Widen/In/";
            //string outFolder = @"/Users/matt.howard/Desktop/Widen/Out/";
            string inFolder = settings.inPath;            
            string outFolder = settings.outPath;
            bool deleteZips = settings.deleteZips;

            Directory.CreateDirectory(inFolder);
            Directory.CreateDirectory(outFolder);

            DirectoryInfo dInfo = new DirectoryInfo(inFolder);
            FileInfo[] files = dInfo.GetFiles("*.zip");
            foreach (FileInfo f in files)
            {
                Run(f.FullName, outFolder, deleteZips);
            }
        }

        /// <summary>
        /// Runs the main process
        /// </summary>
        /// <param name="fileName">The zip file to process</param>
        /// <param name="to">Where to export the pdfs</param>
        /// <param name="deleteZips">Shoudl the zips be deleted?</param>
        private static void Run(string fileName, string to, bool deleteZips)
        {
            try
            {
                Console.WriteLine(DateTime.Now + " | Extracting " + fileName + "...");
                UnzipNew(fileName, Path.Combine(to, Path.GetFileNameWithoutExtension(fileName)));
                Console.WriteLine("-------------------------------------------------------------");
                List<string> files = new List<string>(Directory.GetFiles(Path.Combine(to, Path.GetFileNameWithoutExtension(fileName)), "*.pdf", SearchOption.AllDirectories));
                Directory.CreateDirectory(Path.Combine(to, Path.GetFileNameWithoutExtension(fileName)));

                foreach (string f in files)
                {
                    File.Copy(f, Path.Combine(to, Path.GetFileName(f)), true);
                    //File.Decrypt(Path.Combine(to, Path.GetFileNameWithoutExtension(fileName), Path.GetFileName(f)));
                }
                CleanUp(fileName, to, deleteZips);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// T   Unzips the contents to get the pdfs.
        /// </summary>
        /// <param name="sourceFile">The Zip file</param>
        /// <param name="destination">Where the pdfs need to go.</param>
        private static void UnzipNew(string sourceFile, string destination)
        {
            List<int> linkRemoveList = new List<int>();
            using (ZipArchive archive = ZipFile.OpenRead(sourceFile))
            {
                int zipCount = 0;
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    if (archive.Entries[i].FullName.ToLower().Contains("links"))
                    {
                        linkRemoveList.Add(i);
                    }
                }
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    if ((archive.Entries[i].FullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) && (!linkRemoveList.Contains(i)))
                    {
                        zipCount++;
                    }
                }
                int count = 0;
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    if ((archive.Entries[i].FullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) && (!linkRemoveList.Contains(i)))
                    {
                        if ((!Path.GetFileNameWithoutExtension(archive.Entries[i].FullName).Contains("._")) && (!Path.GetFileNameWithoutExtension(archive.Entries[i].FullName).StartsWith("_")))
                        {
                            count++;
                        }
                    }
                }
                int fileCount = 1;
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    if ((archive.Entries[i].FullName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) && (!linkRemoveList.Contains(i)))
                    {
                        string filename = Path.GetFileName(sourceFile);
                        _ = filename.Replace("-", " ");
                        string destinationPath;

                        if ((!Path.GetFileNameWithoutExtension(archive.Entries[i].FullName).Contains("._")) && (!Path.GetFileNameWithoutExtension(archive.Entries[i].FullName).StartsWith("_")))
                        {
                            if (count < 2)
                            {
                                destinationPath = Path.GetFullPath(Path.Combine(destination, Path.GetFileNameWithoutExtension(filename) + ".pdf"));
                            }
                            else
                            {
                                destinationPath = Path.GetFullPath(Path.Combine(destination, Path.GetFileNameWithoutExtension(filename) + " - " + fileCount.ToString().PadLeft(3, '0') + ".pdf"));
                                fileCount++;
                            }
                            Directory.CreateDirectory(destination);
                            archive.Entries[i].ExtractToFile(destinationPath, true);
                        }

                    }
                }
            }
        }

        /// <summary>
        ///     Delete any left over files
        /// </summary>
        /// <param name="fileName">The zip file name</param>
        /// <param name="to">Where the pdfs were extracted to</param>
        /// <param name="deleteZips">Should the Zips be Deleted?</param>
        private static void CleanUp(string fileName, string to, bool deleteZips)
        {
            try
            {
                bool go = false;
                while (!go)
                {
                    FileInfo dn = new FileInfo(fileName);
                    if (!IsFileLocked(dn))
                    {
                        Directory.Delete(Path.Combine(to, Path.GetFileNameWithoutExtension(fileName)), true);
                        if (deleteZips)
                        {
                            File.Delete(fileName);
                        }
                        go = true;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        ///     Checks to see if a file is locked
        /// </summary>
        /// <param name="file">The file name.</param>
        /// <returns>Returned true or false depending on if the file is locked.</returns>
        private static bool IsFileLocked(FileInfo file)
        {
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }

            //file is not locked
            return false;
        }

        /// <summary>
        ///     Gets the settings from the config.xml file.
        /// </summary>
        /// <returns>Returns the settings.</returns>
        private static Settings GetSettings()
        {
            Settings settings = new Settings();
            XmlDocument doc = new XmlDocument();
            //string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "Config.xml");
            string xmlPath = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location), "Config.xml");
            doc.Load(xmlPath);

            settings.inPath = doc.GetElementsByTagName("inPath")[0].InnerText;
            settings.outPath = doc.GetElementsByTagName("outPath")[0].InnerText;
            settings.deleteZips = Convert.ToBoolean(doc.GetElementsByTagName("deleteZips")[0].InnerText);

            return settings;
        }        
    } 
}
public class Settings
{
    public string inPath { get; set; }
    public string outPath { get; set; }
    public bool deleteZips { get; set; }

}
