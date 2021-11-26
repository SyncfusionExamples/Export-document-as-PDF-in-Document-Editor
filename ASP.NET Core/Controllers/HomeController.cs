using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using PdfExport.Models;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using EJ2DocumentEditor = Syncfusion.EJ2.DocumentEditor;

namespace PdfExport.Controllers
{
    public class HomeController : Controller
    {
        private IHostingEnvironment hostEnvironment;
        public HomeController(IHostingEnvironment environment)
        {
            this.hostEnvironment = environment;
        }

        public IActionResult Index()
        {
            //ViewBag.filePathInfo = GetFilesInfo();
            string fileName = "Non-Disclosure Agreement.docx";
            //string fileName = "Contract.docx";
            string path = this.hostEnvironment.WebRootPath + "\\Files\\" + fileName;
            //Converts the base64 string to byte array

            byte[] byteArray = System.IO.File.ReadAllBytes(path);
            Stream stream = new MemoryStream(byteArray);
            Syncfusion.EJ2.DocumentEditor.WordDocument document = Syncfusion.EJ2.DocumentEditor.WordDocument.Load(stream, GetFormatType(fileName));
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(document);
            document.Dispose();
            stream.Dispose();

            ViewBag.username = "Test User";
            ViewBag.fileName = fileName;
            ViewBag.json = json;
            return View();
        }
        public List<FilesPathInfo> GetFilesInfo()
        {
            string path = hostEnvironment.WebRootPath + "\\Files";
            return ExplorFiles(path);
        }

        // GET: FileExplorer
        public List<FilesPathInfo> ExplorFiles(string fileDirectory)
        {
            List<FilesPathInfo> filesInfo = new List<FilesPathInfo>();
            try
            {
                foreach (string f in Directory.GetFiles(fileDirectory))
                {
                    FilesPathInfo path = new FilesPathInfo();
                    path.text = Path.GetFileName(f);
                    filesInfo.Add(path);
                }
            }
            catch (System.Exception e)
            {
                throw new Exception("error", e);
            }
            return filesInfo;
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
    }
}
