using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;

namespace DocAssemblyPrototype
{
    public enum DocumentType
    {
        Undefined,
        Docx,
        Xlsx
    }

    public class DocumentHelper
    {
        private DocumentType DocumentType { get; set; }

        public DocumentHelper(string filename)
        {
            if (Path.GetExtension(filename) == ".docx")
                DocumentType = DocumentType.Docx;
            else if (Path.GetExtension(filename) == ".xlsx")
                DocumentType = DocumentType.Xlsx;
            else
                DocumentType = DocumentType.Undefined;
        }

        public string DocumentPath
        {
            get {
                if (DocumentType == DocumentType.Docx)
                    return "word/document.xml";
                else if(DocumentType == DocumentType.Xlsx)
                    return "xl/sharedStrings.xml";

                throw new Exception("unknown file type");
            }
        }
    }
}
