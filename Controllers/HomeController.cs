using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.IO.Compression;


using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using System.Text.RegularExpressions;

// For more information on enabling MVC for empty projects, visit http://go.microsoft.com/fwlink/?LinkID=397860

namespace DocAssemblyPrototype.Controllers
{
    [Route("[controller]")]
    public class HomeController : Controller
    {

        private readonly string _pathTemplates;
        private readonly string _pathDocuments;

        public HomeController()
        {
            var pathDocAssemblyPrototype = Path.Combine(Path.GetTempPath(), "DocAssemblyPrototype");
            _pathTemplates = Path.Combine(pathDocAssemblyPrototype, "Templates");
            _pathDocuments = Path.Combine(pathDocAssemblyPrototype, "Documents");

            foreach (var dir in new[] { pathDocAssemblyPrototype, _pathTemplates, _pathDocuments })
            {
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
            }
        }
           
        [Route("GetTemplateList")]
        [Route("templates")]
        [HttpGet]
        public IActionResult GetTemplateList()
        {
            var list = new List<string>();
            Directory.EnumerateFiles(_pathTemplates).ToList().ForEach(x => list.Add(Path.GetFileName(x)));
            return new JsonResult(list);
        }

        [Route("GetTemplateInfo/{filename}")]
        [Route("templates/{filename}")]
        [HttpGet]
        public IActionResult GetTemplateInfo(string filename)
        {
            string xml = null;
            var docHelper = new DocumentHelper(filename);
            xml = DocumentProcessor.GetDocumentXml(Path.Combine(_pathTemplates, filename), docHelper.DocumentPath);
                      
            if (string.IsNullOrWhiteSpace(xml))
            {
                return new JsonResult("Incorrect File Format");
            }

            var list = DocumentProcessor.GetListOfFields(xml);
            var ur = new TemplateInfo { Fields = list.ToDictionary(x => x), TemplateName = filename };

            return new JsonResult(ur);
        }

       
        [Route("UploadTemplate")]
        [Route("templates")]        
        [HttpPost]
        [HttpPut]
        public IActionResult UploadTemplate()            
        {
            Stream uploadFileStream = null;
            string templateName;

            if (this.HttpContext.Request.HasFormContentType)
            {
                uploadFileStream = this.HttpContext.Request.Form.Files[0].OpenReadStream();
                templateName = this.HttpContext.Request.Form.Files[0].FileName;
            }            
            else
            {
                if (this.HttpContext.Request.QueryString.HasValue)
                {
                    templateName = HttpContext.Request.QueryString.Value;
                }
                else
                {
                    var guid = Guid.NewGuid().ToString("N");
                    templateName = string.Format("template_{0}.docx", guid);
                }
                
                uploadFileStream = this.HttpContext.Request.Body;
            }

            string xml = null;
            var docHelper = new DocumentHelper(templateName);

            using (var fileStream = System.IO.File.Create(Path.Combine(_pathTemplates, templateName)))
            {
                uploadFileStream.CopyTo(fileStream);
                xml = DocumentProcessor.GetDocumentXml(fileStream, docHelper.DocumentPath);                
            }

            if (string.IsNullOrWhiteSpace(xml))
            {
                return new JsonResult("Incorrect File Format");
            }

            var list = DocumentProcessor.GetListOfFields(xml);
            var ur = new TemplateInfo { Fields = list.ToDictionary(x => x), TemplateName = templateName };

            return new JsonResult(ur);
        }

        [Route("GenerateDocument")]
        [Route("documents")]
        [HttpPost]
        public IActionResult GenerateDocument([FromBody]TemplateInfo ur)
        {
            var guid = Guid.NewGuid().ToString("N");
            var docHelper = new DocumentHelper(ur.TemplateName);
            var docName = ur.TemplateName.Replace("template_", string.Empty).Replace(".docx", "_" + guid + ".docx").Replace(".xlsx", "_" + guid + ".xlsx");
            using (var fileStream = System.IO.File.Open(Path.Combine(_pathTemplates, ur.TemplateName), FileMode.Open))
            using (var fileStreamDoc = System.IO.File.Open(Path.Combine(_pathDocuments, docName), FileMode.Create))
            {
                fileStream.CopyTo(fileStreamDoc);
                DocumentProcessor.ReplaceFields(fileStreamDoc, ur.Fields, docHelper.DocumentPath);
            }

            return new JsonResult(docName);
        }

        [Route("GetDocument/{filename}")]
        [Route("documents/{filename}")]
        [HttpGet]
        public IActionResult GetDocument(string filename)
        {
            var fileStream = System.IO.File.Open(Path.Combine(_pathDocuments, filename), FileMode.Open);
            var fsr = new FileStreamResult(fileStream, "application/docx");
            fsr.FileDownloadName = filename;
            return fsr;
        }

        [Route("Test")]
        [Route("")]
        [HttpGet]
        public IActionResult Test()
        {
            return View();
        }
    }
}
