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
       
        public IActionResult Index()
        {
          
            return View();
        }
     
    }
}
