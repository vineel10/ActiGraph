﻿using System;
using System.IO;
using System.Linq;

namespace Actigraph.Parser
{
    public static class DirectoryStructure
    {
        public static string ActigraphDataFilesFolderName=string.Empty;
        public static string ActigraphDataFilesProcessedFolderName = string.Empty;
        public static string ActigraphReportsFolderName = string.Empty;
        public static string CurrentDateTime = String.Empty;
        public static int i = 0;

        static DirectoryStructure()
        {
            try
            {

            ActigraphDataFilesFolderName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ActigraphDataFiles");
            ActigraphDataFilesProcessedFolderName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ActigraphDataFilesProcessed");
            ActigraphReportsFolderName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ActigraphReports");
            CurrentDateTime = DateTime.Now.ToShortDateString().Replace('/','-');
            }
            catch (Exception exp)
            {
                throw exp;
            }
        }
        public static void CreateDirectoryStructure()
        {
            try
            {

            // ActigraphDataFilesFolderName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ActigraphDataFiles");

            if (!Directory.Exists(ActigraphDataFilesFolderName))
                 Directory.CreateDirectory(ActigraphDataFilesFolderName);

            //ActigraphDataFilesProcessedFolderName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ActigraphDataFilesProcessed");
            if (!Directory.Exists(ActigraphDataFilesProcessedFolderName))
                Directory.CreateDirectory(ActigraphDataFilesProcessedFolderName);

            //ActigraphReportsFolderName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ActigraphReports");
            if (!Directory.Exists(ActigraphReportsFolderName))
                Directory.CreateDirectory(ActigraphReportsFolderName);

            }
            catch (Exception exp)
            {

                throw exp;
            }

        }

        public static string[] GetStagedFiles()
        {
            try
            {

                var stagedFiles = Directory.GetFiles(ActigraphDataFilesFolderName).Where(f=>
                {
                    var extension = Path.GetExtension(f);
                    return extension != null && extension.ToUpperInvariant().Contains("XLS");
                });
                return stagedFiles.ToArray();
            }
            catch (Exception exp)
            {

                throw exp;
            }
        }

        public static string CreateSubjectFolder(string folderName)
        {
            try
            {
           
            var currentFolderName = Path.Combine(ActigraphReportsFolderName,CurrentDateTime, folderName + "-" + CurrentDateTime);
            if (!Directory.Exists(currentFolderName))
            {
                Directory.CreateDirectory(currentFolderName);
                return currentFolderName;
            }
            return currentFolderName;
            }
            catch (Exception exp)
            {

                throw exp;
            }
        }

        public static void MoveProcessedFile(string processedFile)
        {
            
            string fileName = processedFile;
            try
            {
            string absolutePath = Path.Combine(DirectoryStructure.ActigraphDataFilesFolderName, processedFile);
            File.Copy(absolutePath, Path.Combine(ActigraphDataFilesProcessedFolderName,Path.GetFileNameWithoutExtension(processedFile) + CurrentDateTime + Path.GetExtension(processedFile)),true);
            File.Delete(absolutePath);
            }
            catch (Exception exp)
            {
                if (i < 10)
                {
                    MoveProcessedFile(fileName);
                    i++;
                }
            }
        }
    }
}
