using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO.Compression;
using ExcelUnlock.Properties;

namespace ExcelUnlock
{
    /// <summary>
    /// Single custom class for unlock excel file.
    /// </summary>
       
    class ExcelUnlockClass
    {
        public string pathOrginalExcel { get; set; } = "";
        public string pathTempFolder { get; set; } = System.IO.Path.GetTempPath();
        public string pathTo { get; set; } = "";
        private string fileExtension = ".xlsx";
        private List<string> fileToDelete = new List<string>();

        public EventHandler<ReportArgs> reportEvent;

        public ExcelUnlockClass(string path)
        {
            pathOrginalExcel = path;
        }

        public void process()
        {
            //create path to temp zip file
            string pathTempExcel = Path.Combine(pathTempFolder, Path.GetFileName(pathOrginalExcel));
            string pathTempZipFull = Path.ChangeExtension(pathTempExcel, ".zip");
            string pathExtracted = Path.ChangeExtension(pathTempZipFull, null);

            fileCheck();
            clearTemp(pathExtracted, pathTempZipFull);
            copyToTemp(pathTempZipFull);
            extractedFile(pathTempZipFull, pathExtracted);
            unlockEachExtractedFile(pathExtracted);
            importXmlFilesToZip(pathExtracted, pathTempZipFull);
            copyFromTemp(pathTempZipFull);
            clearTemp(pathExtracted, pathTempZipFull);
            report(Resources.EXCRepFinished);
        }

        private void fileCheck()
        {
            //check if it is right excel file
            report(Resources.EXRepCheck);

            if (string.IsNullOrEmpty(pathOrginalExcel))
            {
                Exception exc = new Exception(Resources.EXExcEmptyPath);
                throw exc;
            }

            if (!File.Exists(pathOrginalExcel))
            {
                Exception exc = new Exception(Resources.EXExcFileNoExist);
                throw exc;
            }

            if (!(Path.GetExtension(pathOrginalExcel)).Equals(fileExtension))
            {
                Exception exc = new Exception($"{Resources.EXExcWrongExtension} {fileExtension}");
                throw exc;
            }

            if (string.IsNullOrEmpty(pathTempFolder))
            {
                Exception exc = new Exception(Resources.EXExcEmptyTempPath);
                throw exc;
            }

            if (!Directory.Exists(pathTempFolder))
            {
                Exception exc = new Exception(Resources.EXExcTempPath);
                throw exc;
            }
        }

        private void clearTemp(string pathExtracted, string pathTempZipFull)
        {
            //remove specific items from temp
            report(Resources.EXReClearTemp);

            if (Directory.Exists(pathExtracted))
            {
                Directory.Delete(pathExtracted, true);
            }
            if (File.Exists(pathTempZipFull))
            {
                File.Delete(pathTempZipFull);
            }
        }

        private void copyToTemp(string pathTempZipFull)
        {
            //create zip file in temp
            report($"{Resources.EXReCopyTemp} ({pathTempFolder}).");

            if (!File.Exists(pathTempZipFull))
            {
                File.Copy(pathOrginalExcel, pathTempZipFull);
            }
            else
            {
                File.Delete(pathTempZipFull);
                File.Copy(pathOrginalExcel, pathTempZipFull);

                //Exception exc = new Exception("Není možné vytvořit dočasný soubor. " + pathTempFull);
                //throw exc;
            }
            fileToDelete.Add(pathTempZipFull);
        }

        private void extractedFile(string pathTempZipFull, string pathExtracted)
        {
            //extract file
            report(Resources.EXReUnpackEx);
            System.IO.Compression.ZipFile.ExtractToDirectory(pathTempZipFull, pathExtracted);
        }

        private void unlockEachExtractedFile(string pathExtracted)
        {
            //edit each sheet file and remove protection node
            report(Resources.EXReUnlockLists);

            string[] pathsToCombine = new string[] { pathExtracted, "xl", "worksheets" };
            string pathToLockedFolder = Path.Combine(pathsToCombine);
            string[] allFiles = Directory.GetFiles(pathToLockedFolder);
            foreach (var file in allFiles)
            {
                editFile(file);
            }
        }

        private void editFile(string pathFile)
        {
            //edit and save file
            report($"{Resources.EXReUnlockList} {Path.GetFileName(pathFile)}.");

            string[] oldText = System.IO.File.ReadAllLines(pathFile);
            List<String> newText = new List<String>();

            foreach (string text in oldText)
            {
                newText.Add(replaceString(text));
            }

            System.IO.File.WriteAllLines(pathFile, newText.ToArray());
        }

        private string replaceString(string textToReplace)
        {
            //remove phrase with sheetProtection

            //string textToReplace = @"neco1 neco2<sheetProtection algorithmName=""SHA - 512"" hashValue =""7KQBa4GBltyULA8DN99CvRWqJqzuIUy9QXFkngOQO5abQ32Voqo5G8ABpwwJunSTk4cOFj + Rd3LR3yhazX0oTA == "" saltValue =""mJqjzpQp9U + GJbObFzSd6A == "" spinCount =""100000"" sheet =""1"" objects =""1"" scenarios =""1"" formatCells=""0"" formatColumns=""0"" formatRows=""0"" sort=""0"" autoFilter=""0""/> neco neco";
            string matchCodeTag = @"\<sheetProtection(.*?)\/>";
            string replaceWith = "";
            string output = Regex.Replace(textToReplace, matchCodeTag, replaceWith);
            return output;
        }

        private void deleteFileInZip(string pathZipFile, string pathInZip)
        {
            //delete (sheet) file in zip 
            report(Resources.EXRePrepComplet);

            using (FileStream zipToOpen = new FileStream(pathZipFile, FileMode.Open))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    foreach (var item in archive.Entries)
                    {
                        if (item.Name.Equals(pathInZip))
                        {
                            item.Delete();
                            break; //needed to break out of the loop
                        }
                    }
                }

            }

        }
        
        private void importXmlFilesToZip(string pathExtracted, string pathZipFile)
        {
            report(Resources.EXReCompleting);
            //import files from specific path to zip file (delete orginal)
            string[] pathFoldersInZip = new string[] { "xl", "worksheets" };

            //get extracted files
            List<string> pathToCombineUnziped = pathFoldersInZip.ToList();
            pathToCombineUnziped.Insert(0, pathExtracted);
            string pathExtractedSheets = Path.Combine(pathToCombineUnziped.ToArray());
            string[] files = Directory.GetFiles(pathExtractedSheets);

            //add files to zip
            foreach (string pathOneFile in files)
            {
                string fileName = Path.GetFileName(pathOneFile).ToString();
                List<string> pathToCombineZip = pathFoldersInZip.ToList();
                pathToCombineZip.Add(fileName);
                string pathInZip = Path.Combine(pathToCombineZip.ToArray());

                deleteFileInZip(pathZipFile, fileName);

                //write to Zip file
                importFileToZipFile(pathOneFile, pathZipFile, pathInZip);
            }
        }

        private void importFileToZipFile(string pathToUnzipedFile, string pathZipFile, string pathFileInsideZip)
        {
            //read file and create them in zip file

            //example
            //string pathFileInsideZip = @"xl\worksheets\sheet1.xml";
            //string pathToUnzipedFile = @"C:\Users\jvaldauf\AppData\Local\Temp\testZamknutehoExcelu\xl\worksheets\";
            //string pathZipFile = @"C:\Users\jvaldauf\AppData\Local\Temp\testZamknutehoExcelu.zip";

            //read file
            string[] lines = System.IO.File.ReadAllLines(pathToUnzipedFile);
            
            //edit text in ziped text file
            using (FileStream zipToOpen = new FileStream(pathZipFile, FileMode.Open))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry = archive.CreateEntry(pathFileInsideZip);
                    using (StreamWriter writer = new StreamWriter(readmeEntry.Open()))
                    {
                        //write lines in ziped text file
                        foreach (string line in lines)
                        {
                            writer.WriteLine(line);
                        }
                    }
                }
            }

        }

        private void copyFromTemp(string pathTempZipFull)
        {
            //copy zip from temp to orginal location (xlsx)
            report(Resources.EXReSendingBack);

            string fileName = Path.GetFileNameWithoutExtension(pathOrginalExcel) + Resources.EXUnlockedTag + fileExtension;
            string pathOrginal = Path.GetDirectoryName(pathOrginalExcel);
            string pathTo      = Path.Combine(pathOrginal, fileName);

            if (!File.Exists(pathTo))
            {
                File.Copy(pathTempZipFull, pathTo);
            }
            else
            {
                File.Delete(pathTo);
                File.Copy(pathTempZipFull, pathTo);
            }
        }

        private void report(string reportText)
        {
            //send back info about progress
            if (reportEvent != null)
            {
                reportEvent(null, new ReportArgs { report = reportText });
            }
        }
    }
}
