using PdfExport.Models;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.DocIO;
using System;
using System.IO;
using EJ2DocumentEditor = Syncfusion.EJ2.DocumentEditor;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;

namespace PdfExport.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DocumenteditorController : ControllerBase
    {
        private IHostingEnvironment hostEnvironment;
        public DocumenteditorController(IHostingEnvironment environment)
        {
            this.hostEnvironment = environment;
        }
        //Import file from client side.
        [Route("Import")]
        public string Import(IFormCollection data)
        {
            if (data.Files.Count == 0)
                return null;
            Stream stream = new MemoryStream();
            IFormFile file = data.Files[0];
            int index = file.FileName.LastIndexOf('.');
            string type = index > -1 && index < file.FileName.Length - 1 ?
                file.FileName.Substring(index) : ".docx";
            file.CopyTo(stream);
            stream.Position = 0;

            EJ2DocumentEditor.WordDocument document = EJ2DocumentEditor.WordDocument.Load(stream, GetFormatType(type.ToLower()));
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            return json;
        }
        //Import documents from web server.
        [Route("ImportFile")]
        public string ImportFile([FromBody]CustomParams param)
        {
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + param.fileName;
            try
            {
                Stream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.ReadWrite);
                Syncfusion.EJ2.DocumentEditor.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(stream, GetFormatType(path));
                string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
                document.Dispose();
                stream.Dispose();
                return json;
            }
            catch
            {
                return "Failure";
            }
        }
        public class SaveParam
        {
            public string content { get; set; }
        }
        [AcceptVerbs("Post")]
        [HttpPost]
        [Route("ExportPdf")]
        public FileStreamResult ExportPdf([FromBody] SaveParameter data)
        {
            // Converts the sfdt to stream
            Stream document = EJ2DocumentEditor.WordDocument.Save(data.Content, EJ2DocumentEditor.FormatType.Docx);
            Syncfusion.DocIO.DLS.WordDocument doc = new Syncfusion.DocIO.DLS.WordDocument(document, Syncfusion.DocIO.FormatType.Docx);
            //Instantiation of DocIORenderer for Word to PDF conversion 
            DocIORenderer render = new DocIORenderer();
            //Converts Word document into PDF document 
            PdfDocument pdfDocument = render.ConvertToPDF(doc);
            Stream stream = new MemoryStream();
            
            //Saves the PDF file
            pdfDocument.Save(stream);
            stream.Position = 0;
            pdfDocument.Close();         
            document.Close();
            return new FileStreamResult(stream, "application/pdf")
            {
                FileDownloadName = data.FileName
            };
        }
        public class SaveParameter
        {
            public string Content { get; set; }
            public string FileName { get; set; }
        }

        [Route("ExportSfdt")]
        public void ExportSfdt([FromBody] SaveParameter data)
        {
            string name = data.FileName;
            //string format = GetFormatType(name);
            if (string.IsNullOrEmpty(name))
            {
                name = "Document1.doc";
            }
            Syncfusion.DocIO.DLS.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Save(data.Content);
            FileStream fileStream = new FileStream(name, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            document.Save(fileStream, Syncfusion.DocIO.FormatType.Docx);
            document.Close();
            fileStream.Close();
        }
        //Save document in web server.
        [Route("Save")]
        public string Save([FromBody]CustomParameter param)
        {
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + param.fileName;
            Byte[] byteArray = Convert.FromBase64String(param.documentData);
            Stream stream = new MemoryStream(byteArray);
            EJ2DocumentEditor.FormatType type = GetFormatType(path);
            try
            {
                FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite);

                if (type != EJ2DocumentEditor.FormatType.Docx)
                {
                    Syncfusion.DocIO.DLS.WordDocument document = new Syncfusion.DocIO.DLS.WordDocument(stream, Syncfusion.DocIO.FormatType.Docx);
                    document.Save(fileStream, GetDocIOFomatType(type));
                    document.Close();
                }
                else
                {
                    stream.Position = 0;
                    stream.CopyTo(fileStream);
                }
                stream.Dispose();
                fileStream.Dispose();
                return "Sucess";
            }
            catch
            {
                Console.WriteLine("err");
                return "Failure";
            }
        }

        internal static EJ2DocumentEditor.FormatType GetFormatType(string fileName)
        {
            int index = fileName.LastIndexOf('.');
            string format = index > -1 && index < fileName.Length - 1 ? fileName.Substring(index + 1) : "";

            if (string.IsNullOrEmpty(format))
                throw new NotSupportedException("EJ2 Document editor does not support this file format.");
            switch (format.ToLower())
            {
                case "dotx":
                case "docx":
                case "docm":
                case "dotm":
                    return EJ2DocumentEditor.FormatType.Docx;
                case "dot":
                case "doc":
                    return EJ2DocumentEditor.FormatType.Doc;
                case "rtf":
                    return EJ2DocumentEditor.FormatType.Rtf;
                case "txt":
                    return EJ2DocumentEditor.FormatType.Txt;
                case "xml":
                    return EJ2DocumentEditor.FormatType.WordML;
                default:
                    throw new NotSupportedException("EJ2 Document editor does not support this file format.");
            }
        }

        internal static Syncfusion.DocIO.FormatType GetDocIOFomatType(EJ2DocumentEditor.FormatType type)
        {
            switch (type)
            {
                case EJ2DocumentEditor.FormatType.Docx:
                    return FormatType.Docx;
                case EJ2DocumentEditor.FormatType.Doc:
                    return FormatType.Doc;
                case EJ2DocumentEditor.FormatType.Rtf:
                    return FormatType.Rtf;
                case EJ2DocumentEditor.FormatType.Txt:
                    return FormatType.Txt;
                case EJ2DocumentEditor.FormatType.WordML:
                    return FormatType.WordML;
                default:
                    throw new NotSupportedException("DocIO does not support this file format.");
            }
        }
    }
}
