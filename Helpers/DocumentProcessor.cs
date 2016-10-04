using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace DocAssemblyPrototype
{
    public class DocumentProcessor
    {
        public static string GetDocumentXml(Stream fileStream, string documentPath)
        {
            using (ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Read))
            {
                var docXml = archive.GetEntry(documentPath);
                using (var sr = new StreamReader(docXml.Open()))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        public static string GetDocumentXml(string filePath, string documentPath)
        {
            using (var fileStream = File.Open(filePath, FileMode.Open))
            {
                return GetDocumentXml(fileStream, documentPath);
            }
        }

        public static void ReplaceFields(Stream fileStream, Dictionary<string, string> fields, string documentPath)
        {
            using (ZipArchive archive = new ZipArchive(fileStream, ZipArchiveMode.Update))
            {
                string xml;
                var docXml = archive.GetEntry(documentPath);

                using (var sr = new StreamReader(docXml.Open(), System.Text.Encoding.UTF8))
                {
                    xml = sr.ReadToEnd();
                }

                fields.ToList().ForEach(keyPair => xml = xml.Replace(keyPair.Key, keyPair.Value));

                docXml.Delete();

                docXml = archive.CreateEntry(documentPath);

                using (var sw = new StreamWriter(docXml.Open(), System.Text.Encoding.UTF8))
                {
                    sw.Write(xml);
                }
            }
        }

        public static List<string> GetListOfFields(string text)
        {
            var rg = new Regex(@"\[[a-zA-Z0-9 :]+\]");
            var matches = rg.Matches(text);
            var list = new List<string>();

            foreach (Match match in matches)
            {
                if (list.Contains(match.Value))
                    continue;

                list.Add(match.Value);
            }
            
            list.Sort();

            return list;
        }
    }
}
