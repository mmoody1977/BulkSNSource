using BulkKBUpload.Domain;
using BulkKBUpload.WebServices;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BulkKBUpload.AppServices
{
    public class KBUpload
    {
        private SN_Service sn_service;
        private string baseSNURL;
        private string snUsername;
        private string snPassword;
        public Logger logger { get; set; }
        public KBUpload(Logger Logger)
        {
            if (String.IsNullOrEmpty(snUsername)) {
                snUsername = getConfigurationFileValue("sn_username");
            } 
            if (String.IsNullOrEmpty(snPassword))
            {
                snPassword = getConfigurationFileValue("sn_password");
            }
            if (String.IsNullOrEmpty(baseSNURL))
            {
                baseSNURL = getConfigurationFileValue("sn_instance_url");
            }
            if (sn_service == null)
            {
                sn_service = new SN_Service(baseSNURL, snUsername, snPassword);
            }
            logger = Logger;
        }

        public async Task BulkUpload(string FolderPath)
        {
            List<Task> uploadTasks = new List<Task>();
            Console.WriteLine("******Directory: " + FolderPath);
            string[] files = Directory.GetFiles(FolderPath,"*.doc");
            foreach (string filePath in files)
            {
                Console.WriteLine("#####################################");
                Console.WriteLine("Uploading file: " + filePath);
                await logger.Log("Uploading file: " + filePath, logType.Info);
                Console.WriteLine("#################");
                Task task = uploadFile(filePath);
                uploadTasks.Add(task);
                //await uploadFile(filePath);
            }
            try
            {
                foreach (string dirPath in Directory.GetDirectories(FolderPath))
                {
                    await BulkUpload(dirPath);
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                await logger.Log(String.Format("Unable to access and upload files in directory {0} | Exception: {1}", FolderPath, ex.Message), logType.Error);
                Console.WriteLine("Unable to access and upload files in directory {0} | Exception: {1}", FolderPath, ex.Message);
            }
            if (uploadTasks.Count() > 0)
            {
                Task.WaitAll(uploadTasks.ToArray());
            }
        }

        private async Task uploadFile(string FilePath)
        {
            //bool fileUploaded = true;

            // Get File Extension
            string fileExtension = "." + FilePath.Split('.')[1];
            // Get File Name
            string[] splitFile = FilePath.Split('\\');
            string fileName = splitFile[splitFile.Length - 1];
            string splitTarget = "\\" + fileName;
            //Console.WriteLine("splitTarget: " + splitTarget);
            
            string targetPath = FilePath.Replace(splitTarget, "");
            //Console.WriteLine("targetPath: " + targetPath);
            //.ReadLine();
            // Convert File to Html
            docConvert DocConvert = new docConvert(sn_service, logger);
            
            string htmlv2 = docConvert.ConvertToHtml(FilePath, targetPath);
           // string html = docConvert.convertToHTML(FilePath);

            // Move CSS Styling inline
            htmlv2 = docConvert.MoveCssInline(htmlv2);

            //string html = docConvert.ConvertDocxToHTML(FilePath);
            // Instantiate Submission
            Submission fileSubmission = new Submission()
            {            
                html = htmlv2,    
                short_description = FilePath
            };
            // Create kb_submission record and set the sys_id
            fileSubmission.sys_id = await sn_service.CreateKBSubmission(fileSubmission);
            // Send Images to ServiceNow and update fileSubmission.html | Strip out and use HTML Body only
            
            //DocConvert.extractImages(ref fileSubmission); 
            DocConvert.sendImages(ref fileSubmission, targetPath);
            // 
            //using (var sw = File.CreateText(@"c:\kb_upload\output.txt"))
            //{
            //    sw.Write(fileSubmission.html);
            //}

            // Update HTML field on kb_submission record in SN
            sn_service.UpdateKBSubmission(fileSubmission);
            // Upload source file as attachment on Submission record
            sn_service.uploadFileAsAttachment(FilePath, fileSubmission);


            //Console.WriteLine(html);
            //return fileUploaded;
        }

        private string getConfigurationFileValue(string key)
        {
            string value;
            value = ConfigurationManager.AppSettings.Get(key);
            return value;
        }
    }
}
