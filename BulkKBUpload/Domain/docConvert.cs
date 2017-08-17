using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mammoth;
using HtmlAgilityPack;
using BulkKBUpload.WebServices;
using OpenXmlPowerTools;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using System.Drawing.Imaging;
using BulkKBUpload.AppServices;

namespace BulkKBUpload.Domain
{
    public class docConvert
    {
        private HtmlDocument htmlDoc { get; set; }
        private SN_Service sn_service { get; set; }
        public Logger logger { get; set; }

        public docConvert(SN_Service SN_Service, Logger Logger)
        {
            if (htmlDoc == null)
            {
                htmlDoc = new HtmlDocument();
            }
            if (sn_service == null)
            {
                sn_service = SN_Service;
            }
            logger = Logger;
        }

        public static string convertToHTML(string FilePath)
        {
            string html = "";
            if (FilePath.Contains(".docx")) // FilePath.Contains(".doc") || 
            {
                var converter = new DocumentConverter();
                var result = converter.ConvertToHtml(FilePath);
                html = result.Value;
                var warnings = result.Warnings;
                if (warnings.Count > 0)
                {
                    foreach (var warning in warnings)
                    {
                        Console.WriteLine("docConvert warning: {0}", warning);
                    }
                }
            }
            return html;
        }

        public void sendImages(ref Submission submission, string FolderPath)
        {
            try
            {
                htmlDoc.LoadHtml(submission.html);
                var imageNodes = htmlDoc.DocumentNode.Descendants("img");
                foreach (var imgNode in imageNodes)
                {
                    var srcAttribs = imgNode.Attributes.AttributesWithName("src");
                    if (srcAttribs.Count() == 1)
                    {
                        // Get Image File Path
                        string imagePath = "";
                        imagePath = FolderPath + "\\" + srcAttribs.First().Value.ToString().Replace('/', '\\');
                        imagePath = imagePath.Replace("%20", " ");
                        Task<List<string>> imageLink = sn_service.uploadImageToSN(imagePath, submission.sys_id);
                        srcAttribs.First().Remove();
                        imgNode.Attributes.Add("src", (string)imageLink.Result[1]);
                       // Console.WriteLine("Image Path: " + imagePath);
                        //Console.ReadLine();
                    }
                }
                submission.html = htmlDoc.DocumentNode.SelectSingleNode("//body").OuterHtml;

                //submission.html = htmlDoc.DocumentNode.OuterHtml;
            }
            catch (Exception ex)
            {
                logger.Log(String.Format("Error sendImages: {0}", ex.Message), logType.Error);
                Console.WriteLine("Error sendImages: {0}", ex.Message);
            }
        }

        public void extractImages(ref Submission submission)
        {
            try
            {
                htmlDoc.LoadHtml(submission.html);
                var imageNodes = htmlDoc.DocumentNode.Descendants("img");
                var nodeId = 0;
                foreach (var imgNode in imageNodes)
                {
                    var srcAttribs = imgNode.Attributes.AttributesWithName("src");
                    if (srcAttribs.Count() == 1)
                    {
                        // Generate Image File Name
                        var imageFileName = "imgFile" + nodeId.ToString();
                        // Get srcAttribute value into parts
                        string[] srcParts = srcAttribs.First().Value.ToString().Split(',');
                        // Get base64 image string
                        string imageBase64 = srcParts[1];
                        // Get contentType
                        string contentType = srcParts[0].Split(':')[1].Split(';')[0]; // "data:image/png;base64,iVBOR"
                        // Upload File and return imageLink
                        Task<List<string>> imageLink = sn_service.uploadImageToSN(imageBase64, imageFileName, contentType, submission.sys_id);
                        srcAttribs.First().Remove();
                        imgNode.Attributes.Add("src", (string)imageLink.Result[0]);
                    }
                    nodeId++;
                }
                submission.html = htmlDoc.DocumentNode.OuterHtml;

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error extractImages: {0}", ex.Message);
            }
        }

        public static string ConvertToHtml(string file, string outputDirectory)
        {
            var fi = new FileInfo(file);
           // Console.WriteLine(fi.Name);
            byte[] byteArray = File.ReadAllBytes(fi.FullName);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
                    if (outputDirectory != null && outputDirectory != string.Empty)
                    {
                        DirectoryInfo di = new DirectoryInfo(outputDirectory);
                        if (!di.Exists)
                        {
                            throw new OpenXmlPowerToolsException("Output directory does not exist");
                        }
                        destFileName = new FileInfo(Path.Combine(di.FullName, destFileName.Name));
                    }
                    var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    int imageCounter = 0;

                    var pageTitle = fi.FullName;
                    var part = wDoc.CoreFilePropertiesPart;
                    if (part != null)
                    {
                        pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
                    }

                    // TODO: Determine max-width from size of content area.
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                        PageTitle = pageTitle,
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            ImageFormat imageFormat = null;
                            if (extension == "png")
                                imageFormat = ImageFormat.Png;
                            else if (extension == "gif")
                                imageFormat = ImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = ImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = ImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                // Convert tiff to gif.
                                extension = "gif";
                                imageFormat = ImageFormat.Gif;
                            }
                            else if (extension == "x-wmf")
                            {
                                extension = "wmf";
                                imageFormat = ImageFormat.Wmf;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            string imageFileName = imageDirectoryName + "/image" +
                                imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.Bitmap.Save(imageFileName, imageFormat);
                            }
                            catch (System.Runtime.InteropServices.ExternalException)
                            {
                                return null;
                            }
                            string imageSource = localDirInfo.Name + "/image" +
                                imageCounter.ToString() + "." + extension;

                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageSource),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement htmlElement = HtmlConverter.ConvertToHtml(wDoc, settings);

                    // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
                    // we are using HTML5.
                    var html = new XDocument(
                        new XDocumentType("html", null, null, null),
                        htmlElement);

                    // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
                    // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                    // for detailed explanation.
                    //
                    // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                    // must do it correctly, or entities will not be serialized properly.

                    var htmlString = html.ToString(SaveOptions.DisableFormatting);
                    //File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                    return htmlString;
                }
            }
        }

        public static string MoveCssInline(string html)
        {
            var result = PreMailer.Net.PreMailer.MoveCssInline(html);
            if (result.Warnings.Count > 0)
            {
                Console.WriteLine("MoveCssInline warnings::");
                foreach (var warning in result.Warnings)
                {
                    
                    Console.WriteLine(warning);
                }
                Console.WriteLine("End MoveCSSInline warnings::");
            }
            return result.Html;
        }

    }
}
