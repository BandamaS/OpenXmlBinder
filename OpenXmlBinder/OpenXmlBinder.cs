using DocumentFormat.OpenXml.Packaging;
using OpenXmlBinder.Exceptions;
using System.Reflection;
using System.Text.RegularExpressions;

namespace OpenXmlBinder
{
    public class OpenXmlBinder : IDisposable
    {
        /// <summary>
        /// String content of the word file
        /// </summary>
        private string TemplateContent { get; set; }

        private MemoryStream MemStream { get; set; }

        private WordprocessingDocument WordDocument { get; set; }

        private Dictionary<string, string?> Variables { get; set; } = new Dictionary<string, string?>();

        public OpenXmlBinder(string filename)
        {
            byte[] byteArray = File.ReadAllBytes(filename);

            MemStream = new MemoryStream();
            MemStream.Write(byteArray);


            WordDocument =  WordprocessingDocument.Open(MemStream, true);

            if (WordDocument.MainDocumentPart == null)
            {
                throw new MainDocumentPartNullException();
            }
            else
            {
                using StreamReader contentStream = new StreamReader(WordDocument.MainDocumentPart.GetStream());


                TemplateContent = contentStream.ReadToEnd();
            }
        }

        /// <summary>
        /// Apply the variables passed on the template and returns a memory stream containing the new word file
        /// </summary>
        /// <returns></returns>
        public Stream Generate()
        {
            foreach(var entry in Variables)
            {
                TemplateContent = Regex.Replace(TemplateContent, entry.Key, entry.Value ?? "");

            }

            using (StreamWriter sw = new StreamWriter(WordDocument.MainDocumentPart!.GetStream(FileMode.Create)))
            {

                sw.Write(TemplateContent);
            }

            return MemStream;
        }

        
        /// <summary>
        /// Adds a variable used in the generate step to fill the template
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public void AddVariable(string key, string? value)
        {
            Variables[key] = value??"";
        }

        /// <summary>
        /// Adds a variable from an object (can be nested)
        /// </summary>
        /// <param name="obj"></param>
        public void AddVariable(object obj)
        {
            var parser = new ObjectParser();

            parser.ParseObject(obj);

            foreach (var entry in parser.parseResult)
            {
                AddVariable(entry.Key, entry.Value);
            }

        }

        public void Dispose()
        {
            WordDocument.Dispose();
            MemStream.Dispose();
        }
    }
}