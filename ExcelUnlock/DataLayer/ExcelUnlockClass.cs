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
    /// Custom class for unlock excel file.
    /// </summary>
    class ExcelUnlockClass
    {
        public EventHandler<ReportArgs> reportEvent;                                //sending back report of current action
        public string pathOrginalExcel { get; set; } = "";                          //file path for orginal locked excel
        public string pathTempFolder { get; set; } = System.IO.Path.GetTempPath();  //folder path to place for extraction
        private string fileExtension = ".xlsx";                                     //extension of excel file
        private string pathTempZipFull;                                             //file path for zip in temp
        private string pathExtracted;                                               //folder path in temp for extracted excel
        private string pathToFinalExcel;                                            //final location of unlocked file

        public ExcelUnlockClass(string path)
        {
            pathOrginalExcel = path;
        }

        /// <summary>
        /// create all inherid paths for temp and final location
        /// </summary>
        private void sortAllPaths()
        {
            //create path to temp zip file
            string pathTempExcel = Path.Combine(pathTempFolder, Path.GetFileName(pathOrginalExcel));
            pathTempZipFull = Path.ChangeExtension(pathTempExcel, ".zip");
            pathExtracted = Path.ChangeExtension(pathTempZipFull, null);

            //create final path
            string finalExcelName = Path.GetFileNameWithoutExtension(pathOrginalExcel) + Resources.EXUnlockedTag + fileExtension;
            string pathOrginal = Path.GetDirectoryName(pathOrginalExcel);
            pathToFinalExcel = Path.Combine(pathOrginal, finalExcelName);
        }

        /// <summary>
        /// execute all steps to unlock excel file
        /// </summary>
        public void process()
        {
            sortAllPaths();
            fileCheck();
            clearTemp(pathExtracted, pathTempZipFull);
            copyToTemp(pathTempZipFull);
            extractedFile(pathTempZipFull, pathExtracted);
            unlockEachExtractedFile(pathExtracted);
            importXmlFilesToZip(pathExtracted, pathTempZipFull);
            copyFromTemp(pathTempZipFull, pathToFinalExcel);
            clearTemp(pathExtracted, pathTempZipFull);
            report(Resources.EXCRepFinished);
        }

        /// <summary>
        /// check paths
        /// </summary>
        private void fileCheck()
        {
            //check if it is right excel file
            report(Resources.EXRepCheck);

            //empty input path
            if (string.IsNullOrEmpty(pathOrginalExcel))
            {
                Exception exc = new Exception(Resources.EXExcEmptyPath);
                throw exc;
            }

            //non-existent input file
            if (!File.Exists(pathOrginalExcel))
            {
                Exception exc = new Exception(Resources.EXExcFileNoExist);
                throw exc;
            }

            //wrong extension file
            if (!(Path.GetExtension(pathOrginalExcel)).Equals(fileExtension))
            {
                Exception exc = new Exception($"{Resources.EXExcWrongExtension} {fileExtension}");
                throw exc;
            }

            //empty path to temp
            if (string.IsNullOrEmpty(pathTempFolder))
            {
                Exception exc = new Exception(Resources.EXExcEmptyTempPath);
                throw exc;
            }

            //non-existent path to temp
            if (!Directory.Exists(pathTempFolder))
            {
                Exception exc = new Exception(Resources.EXExcTempPath);
                throw exc;
            }
        }

        /// <summary>
        /// remove specific folder and file
        /// </summary>
        /// <param name="pathExtracted">folder to delete</param>
        /// <param name="pathTempZipFullLocal">file to delete</param>
        private void clearTemp(string pathExtracted, string pathTempZipFullLocal)
        {
            //remove specific items from temp
            report(Resources.EXReClearTemp);

            if (Directory.Exists(pathExtracted))
            {
                Directory.Delete(pathExtracted, true);
            }
            if (File.Exists(pathTempZipFullLocal))
            {
                File.Delete(pathTempZipFullLocal);
            }
        }

        /// <summary>
        /// copy file to temp
        /// </summary>
        /// <param name="pathTempZipFullLocal">name of file in temp</param>
        private void copyToTemp(string pathTempZipFullLocal)
        {
            report($"{Resources.EXReCopyTemp} ({pathTempFolder}).");

            if (!File.Exists(pathTempZipFullLocal))
            {
                File.Copy(pathOrginalExcel, pathTempZipFullLocal);
            }
            else
            {
                Exception exc = new Exception( Resources.EXExcFullTemp + pathTempZipFullLocal);
                throw exc;
            }
        }

        /// <summary>
        /// extract file to folder
        /// </summary>
        /// <param name="pathTempZipFullLocal">zip file</param>
        /// <param name="pathExtractedLocal">extracted folder</param>
        private void extractedFile(string pathTempZipFullLocal, string pathExtractedLocal)
        {
            report(Resources.EXReUnpackEx);
            System.IO.Compression.ZipFile.ExtractToDirectory(pathTempZipFullLocal, pathExtractedLocal);
        }

        /// <summary>
        /// Go through all sheet files and unlock them
        /// </summary>
        /// <param name="pathExtractedLocal">extracted folder with structure of excel file</param>
        private void unlockEachExtractedFile(string pathExtractedLocal)
        {
            report(Resources.EXReUnlockLists);

            string[] pathsToCombine = new string[] { pathExtractedLocal, "xl", "worksheets" }; //folder with sheets files
            string pathToLockedFolder = Path.Combine(pathsToCombine);
            string[] allFiles = Directory.GetFiles(pathToLockedFolder);
            foreach (var file in allFiles)
            {
                //edit sheet file and remove protection node
                editFile(file);
            }
        }

        /// <summary>
        /// recreate sheet file and remove protection
        /// </summary>
        /// <param name="pathFile">specific excel sheet file</param>
        private void editFile(string pathFile)
        {
            report($"{Resources.EXReUnlockList} {Path.GetFileName(pathFile)}.");
            
            string[] oldText = System.IO.File.ReadAllLines(pathFile); 
            List<String> newText = new List<String>();
            foreach (string text in oldText)
            {
                newText.Add(replaceString(text));
            }
            System.IO.File.WriteAllLines(pathFile, newText.ToArray()); //save edited file
        }

        /// <summary>
        /// remove sheet Protection from string
        /// </summary>
        /// <param name="textToReplace">string with potencial phrase</param>
        /// <returns></returns>
        private string replaceString(string textToReplace)
        {
            //test line of locked sheet
            //string textToReplace = @"neco1 neco2<sheetProtection algorithmName=""SHA - 512"" hashValue =""7KQBa4GBltyULA8DN99CvRWqJqzuIUy9QXFkngOQO5abQ32Voqo5G8ABpwwJunSTk4cOFj + Rd3LR3yhazX0oTA == "" saltValue =""mJqjzpQp9U + GJbObFzSd6A == "" spinCount =""100000"" sheet =""1"" objects =""1"" scenarios =""1"" formatCells=""0"" formatColumns=""0"" formatRows=""0"" sort=""0"" autoFilter=""0""/> neco neco";
            string matchCodeTag = @"\<sheetProtection(.*?)\/>";
            string replaceWith = "";
            string output = Regex.Replace(textToReplace, matchCodeTag, replaceWith);
            return output;
        }

        /// <summary>
        /// add all edited sheets back to zip file
        /// </summary>
        /// <param name="pathExtractedLocal">folder with excel structure</param>
        /// <param name="pathZipFile">orginal zip file</param>
        private void importXmlFilesToZip(string pathExtractedLocal, string pathZipFile)
        {
            report(Resources.EXReCompleting);

            string[] pathFoldersInZip = new string[] { "xl", "worksheets" };

            //get extracted files
            List<string> pathToCombineUnziped = pathFoldersInZip.ToList();
            pathToCombineUnziped.Insert(0, pathExtractedLocal);
            string pathExtractedSheets = Path.Combine(pathToCombineUnziped.ToArray());
            string[] files = Directory.GetFiles(pathExtractedSheets);

            //add files to zip
            foreach (string pathOneFile in files)
            {
                string fileName = Path.GetFileName(pathOneFile).ToString();
                List<string> pathToCombineZip = pathFoldersInZip.ToList();
                pathToCombineZip.Add(fileName);
                string pathInZip = Path.Combine(pathToCombineZip.ToArray());
                //delete sheet in zip
                deleteFileInZip(pathZipFile, fileName);
                //add file to Zip file
                importFileToZipFile(pathOneFile, pathZipFile, pathInZip);
            }
        }

        /// <summary>
        /// remove sheet file from zip file
        /// </summary>
        /// <param name="pathZipFile">zip file</param>
        /// <param name="pathInZip">path in zip file</param>
        private void deleteFileInZip(string pathZipFile, string pathInZip)
        {
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

        /// <summary>
        /// read sheet file and create them in zip file
        /// </summary>
        /// <param name="pathToUnzipedFile">path to sheeet file</param>
        /// <param name="pathZipFile">path to zip file</param>
        /// <param name="pathFileInsideZip">path in zip file</param>
        private void importFileToZipFile(string pathToUnzipedFile, string pathZipFile, string pathFileInsideZip)
        {
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

        /// <summary>
        /// copy editet file to orginal destination (zip to .xlsx)
        /// </summary>
        /// <param name="pathTempZipFullLocal"></param>
        private void copyFromTemp(string pathTempZipFullLocal, string pathToFinalExcelLocal)
        {
            //copy zip from temp to orginal location (xlsx)
            report(Resources.EXReSendingBack);

            if (!File.Exists(pathToFinalExcelLocal))
            {
                File.Copy(pathTempZipFullLocal, pathToFinalExcelLocal);
            }
            else
            {
                File.Delete(pathToFinalExcelLocal);
                File.Copy(pathTempZipFullLocal, pathToFinalExcelLocal);
            }
        }

        /// <summary>
        /// sends back progress information
        /// </summary>
        /// <param name="reportText">text about progress</param>
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
