using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using BulkKBUpload.AppServices;
using System.IO;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using BulkKBUpload.Domain;

namespace BulkKBUpload.WebServices
{
    public class SN_Service
    {
        private string baseURL;
        private string eccInsert = "/ecc_queue.do?JSONv2&sysparm_action=insert";
        private string submissionInsert = "/kb_submission.do?JSONv2&sysparm_action=insert";
        private string attachQuery = "/sys_attachment.do?JSONv2&sysparm_action=getRecords&sysparm_query=file_name=";
        private AuthenticationHeaderValue authHeader;
        public SN_Service(string BaseURL, string UserName, string Password)
        {
            if (String.IsNullOrEmpty(baseURL))
            {
                baseURL = BaseURL;
            }
            if (authHeader == null)
            {
                byte[] authByteArray = Encoding.ASCII.GetBytes(UserName + ":" + Password);
                authHeader = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(authByteArray));
            }
        }

        public async Task<string> CreateKBSubmission(Submission submission)
        {
            string sys_id = "";
            // Create a Submission record and get sys_id
            try
            {
                var jsonObject = new
                {
                    short_description = submission.short_description
                };
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(baseURL);
                    client.DefaultRequestHeaders.Authorization = authHeader;
                    // Execute POST and get response
                    
                    HttpResponseMessage response = await client.PostAsJsonAsync(submissionInsert, jsonObject);
                    // Get response content
                    string content = await response.Content.ReadAsStringAsync();
                    if (content.Contains("HTTP Status 401"))
                    {
                        Console.WriteLine("Error: Unable to create Submission \"{0}\" in ServiceNow. Incorrect Username or Password", submission.short_description);
                    }
                    else if (content.Contains("ServiceNow - Error report"))
                    {
                        Console.WriteLine("Error: Unable to create Submission \"{0}\" in ServiceNow. {1}", submission.short_description, content);
                    }
                    else if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Error: Unable to creat Submission \"{0}\" in ServiceNow. {1}:}{2}", submission.short_description, response.StatusCode, response.ReasonPhrase);
                    }
                    else
                    {
                        //Console.WriteLine("Submission Create response.content: " + content);
                        dynamic jsonResult = JValue.Parse(content);
                        if (jsonResult.records[0].sys_id != "")
                        {
                            sys_id = jsonResult.records[0].sys_id;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in CreateKBSubmission: {0}" + ex.Message);
            }
            return sys_id;
        }

        public async void UpdateKBSubmission(Submission submission)
        {
            // Create a Submission record and get sys_id
            try
            {
                var jsonObject = new
                {

                    text = submission.html
                };
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(baseURL);
                    client.DefaultRequestHeaders.Authorization = authHeader;
                    // Execute POST and get response
                    string submissionUpdate = baseURL + "/kb_submission.do?JSONv2&sysparm_query=sys_id=" + submission.sys_id + "&sysparm_action=update";
                    HttpResponseMessage response = await client.PostAsJsonAsync(submissionUpdate, jsonObject);
                    // Get response content
                    string content = await response.Content.ReadAsStringAsync();
                    if (content.Contains("HTTP Status 401"))
                    {
                        Console.WriteLine("Error: Unable to update Submission \"{0}\" in ServiceNow. Incorrect Username or Password", submission.short_description);
                    }
                    else if (content.Contains("ServiceNow - Error report"))
                    {
                        Console.WriteLine("Error: Unable to update Submission \"{0}\" in ServiceNow. {1}", submission.short_description, content);
                    }
                    else if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Error: Unable to update Submission \"{0}\" in ServiceNow. {1}:}{2}", submission.short_description, response.StatusCode, response.ReasonPhrase);
                    }
                    else
                    {
                        //Console.WriteLine("Submission Update response.content: " + content);
                        //dynamic jsonResult = JValue.Parse(content);
                        //if (jsonResult.records[0].sys_id != "")
                        //{
                        //    sys_id = jsonResult.records[0].sys_id;
                        //}
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error in UpdateKBSubmission: {0}" + ex.Message);
            }
        }

        public async Task<List<string>> uploadImageToSN(string FilePath, string ArticleSid)
        {
            List<string> linkUrls = new List<string>();
            // Get File Extension
            string fileExtension = "." + FilePath.Split('.')[1];
            // Get File Name
            string[] splitFile = FilePath.Split('\\');
            string fileName = splitFile[splitFile.Length - 1];
            linkUrls.Add(fileName);
            // Get File Content Type
            string contentType = MimeTypeMap.GetMimeType(fileExtension);
            // Get File Data
            byte[] fileBinaryData = File.ReadAllBytes(FilePath);
            // Convert File Data to Base64
            string base64 = Convert.ToBase64String(fileBinaryData, 0, fileBinaryData.Length);

            // Upload the file to ServiceNow
            try
            {
                bool imageInsertSuccessful = false;
                var jsonObject = new
                {
                    agent = "AttachmentCreator",
                    topic = "AttachmentCreator",
                    name = fileName + ":" + contentType,
                    source = "kb_submission:" + ArticleSid,
                    payload = base64
                };
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(baseURL);
                    client.DefaultRequestHeaders.Authorization = authHeader;
                    // Execute POST and get respose
                    HttpResponseMessage response = await client.PostAsJsonAsync(eccInsert, jsonObject);
                    // Get response content
                    string content = await response.Content.ReadAsStringAsync();
                    if (content.Contains("HTTP Status 401"))
                    {
                        Console.WriteLine("Error: Unable to upload image file {0} to ServiceNow. Incorrect Username or Password", fileName);
                    }
                    else if (content.Contains("ServiceNow - Error report"))
                    {
                        Console.WriteLine("Error: Unable to upload image file {0} to ServiceNow. {1}", fileName, content);
                    }
                    else if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Error: Unable to upload image file {0} to ServiceNow. {1}:}{2}", fileName, response.StatusCode, response.ReasonPhrase);
                    }
                    else
                    {
                        imageInsertSuccessful = true;
                        //Console.WriteLine("response.content: " + content);
                    }
                }

                // Get the sys_attachment sys_id
                if (imageInsertSuccessful)
                {
                    string sid = "";
                    using (var client = new HttpClient())
                    {
                        client.BaseAddress = new Uri(baseURL);
                        client.DefaultRequestHeaders.Authorization = authHeader;
                        // Execute Query and get response
                        HttpResponseMessage response = await client.GetAsync(attachQuery + fileName + "^table_sys_id=" + ArticleSid);
                        // Get response content
                        string content = await response.Content.ReadAsStringAsync();
                        if (content.Contains("HTTP Status 401"))
                        {
                            Console.WriteLine("Error: Unable to query for image file {0} in ServiceNow. Incorrect Username or Password", fileName);
                        }
                        else if (content.Contains("ServiceNow - Error report"))
                        {
                            Console.WriteLine("Error: Unable to query for image file {0} in ServiceNow. {1}", fileName, content);
                        }
                        else if (!response.IsSuccessStatusCode)
                        {
                            Console.WriteLine("Error: Unable to query for image file {0} in ServiceNow. {1}:}{2}", fileName, response.StatusCode, response.ReasonPhrase);
                        }
                        else
                        {
                            dynamic jsonResult = JValue.Parse(content);
                            if (jsonResult.records[0].sys_id != "")
                            {
                                sid = jsonResult.records[0].sys_id;
                                linkUrls.Add("/sys_attachment.do?sysparm_referring_url=tear_off&view=true&sys_id=" + sid);
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occured during uploadImageToSN: {0}", ex.Message);
            }

            return linkUrls;
        }

        public async Task<List<string>> uploadImageToSN(string FileBase64, string FileName, string ContentType, string ArticleSid)
        {
            List<string> linkUrls = new List<string>();
            //// Get File Extension
            //string fileExtension = "." + FilePath.Split('.')[1];
            //// Get File Name
            //string[] splitFile = FilePath.Split('\\');
            //string fileName = splitFile[splitFile.Length - 1];
            ////linkUrls.Add(fileName);
            //// Get File Content Type
            //string contentType = MimeTypeMap.GetMimeType(fileExtension);
            //// Get File Data
            //byte[] fileBinaryData = File.ReadAllBytes(FilePath);
            //// Convert File Data to Base64
            //string base64 = Convert.ToBase64String(fileBinaryData, 0, fileBinaryData.Length);

            string fileName = FileName;
            string base64 = FileBase64;
            string contentType = ContentType;

            // Upload the file to ServiceNow
            try
            {
                bool imageInsertSuccessful = false;
                var jsonObject = new
                {
                    agent = "AttachmentCreator",
                    topic = "AttachmentCreator",
                    name = fileName + ":" + contentType,
                    source = "kb_submission:" + ArticleSid,
                    payload = base64
                };
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(baseURL);
                    client.DefaultRequestHeaders.Authorization = authHeader;
                    // Execute POST and get respose
                    HttpResponseMessage response = await client.PostAsJsonAsync(eccInsert, jsonObject);
                    // Get response content
                    string content = await response.Content.ReadAsStringAsync();
                    if (content.Contains("HTTP Status 401"))
                    {
                        Console.WriteLine("Error: Unable to upload image file {0} to ServiceNow. Incorrect Username or Password", fileName);
                    }
                    else if (content.Contains("ServiceNow - Error report"))
                    {
                        Console.WriteLine("Error: Unable to upload image file {0} to ServiceNow. {1}", fileName, content);
                    }
                    else if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Error: Unable to upload image file {0} to ServiceNow. {1}:}{2}", fileName, response.StatusCode, response.ReasonPhrase);
                    }
                    else
                    {
                        imageInsertSuccessful = true;
                        //Console.WriteLine("response.content: " + content);
                    }
                }

                // Get the sys_attachment sys_id
                if (imageInsertSuccessful)
                {
                    string sid = "";
                    using (var client = new HttpClient())
                    {
                        client.BaseAddress = new Uri(baseURL);
                        client.DefaultRequestHeaders.Authorization = authHeader;
                        // Execute Query and get response
                        HttpResponseMessage response = await client.GetAsync(attachQuery + fileName + "^table_sys_id=" + ArticleSid);
                        // Get response content
                        string content = await response.Content.ReadAsStringAsync();
                        if (content.Contains("HTTP Status 401"))
                        {
                            Console.WriteLine("Error: Unable to query for image file {0} in ServiceNow. Incorrect Username or Password", fileName);
                        }
                        else if (content.Contains("ServiceNow - Error report"))
                        {
                            Console.WriteLine("Error: Unable to query for image file {0} in ServiceNow. {1}", fileName, content);
                        }
                        else if (!response.IsSuccessStatusCode)
                        {
                            Console.WriteLine("Error: Unable to query for image file {0} in ServiceNow. {1}:}{2}", fileName, response.StatusCode, response.ReasonPhrase);
                        }
                        else
                        {
                            dynamic jsonResult = JValue.Parse(content);
                            if (jsonResult.records[0].sys_id != "")
                            {
                                sid = jsonResult.records[0].sys_id;
                                linkUrls.Add("/sys_attachment.do?sysparm_referring_url=tear_off&view=true&sys_id=" + sid);
                            }
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occured during uploadImageToSN: {0}", ex.Message);
            }

            return linkUrls;
        }

        public async void uploadFileAsAttachment(string FilePath, Submission submission)
        {
            // Get File Extension
            string fileExtension = "." + FilePath.Split('.')[1];
            // Get File Name
            string[] splitFile = FilePath.Split('\\');
            string fileName = splitFile[splitFile.Length - 1];
            //linkUrls.Add(fileName);
            // Get File Content Type
            string contentType = MimeTypeMap.GetMimeType(fileExtension);
            // Get File Data
            byte[] fileBinaryData = File.ReadAllBytes(FilePath);
            // Convert File Data to Base64
            string base64 = Convert.ToBase64String(fileBinaryData, 0, fileBinaryData.Length);


            // Upload the file to ServiceNow
            try
            {
                var jsonObject = new
                {
                    agent = "AttachmentCreator",
                    topic = "AttachmentCreator",
                    name = fileName + ":" + contentType,
                    source = "kb_submission:" + submission.sys_id,
                    payload = base64
                };
                using (var client = new HttpClient())
                {
                    client.BaseAddress = new Uri(baseURL);
                    client.DefaultRequestHeaders.Authorization = authHeader;
                    // Execute POST and get respose
                    HttpResponseMessage response = await client.PostAsJsonAsync(eccInsert, jsonObject);
                    // Get response content
                    string content = await response.Content.ReadAsStringAsync();
                    if (content.Contains("HTTP Status 401"))
                    {
                        Console.WriteLine("Error: Unable to upload source file {0} to ServiceNow. Incorrect Username or Password", fileName);
                    }
                    else if (content.Contains("ServiceNow - Error report"))
                    {
                        Console.WriteLine("Error: Unable to upload source file {0} to ServiceNow. {1}", fileName, content);
                    }
                    else if (!response.IsSuccessStatusCode)
                    {
                        Console.WriteLine("Error: Unable to upload source file {0} to ServiceNow. {1}:}{2}", fileName, response.StatusCode, response.ReasonPhrase);
                    }
                    else
                    {
                       // Console.WriteLine("Source file upload success response.content: " + content);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occured during uploadFileAsAttachment: {0}", ex.Message);
            }

        }

    }
}
