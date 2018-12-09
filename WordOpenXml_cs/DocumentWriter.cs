using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordOpenXml_cs
{
    /// <summary>
    /// Code by Karen Payne MVP along with assistance
    /// from various forum post this class has been glued
    /// together.
    /// </summary>
    public class DocumentWriter : IDisposable
    {
        private MemoryStream _memoryStream;
        /// <summary>
        /// Represents the document to work on
        /// </summary>
        private WordprocessingDocument _document;
        /// <summary>
        /// Create a new document
        /// </summary>
        public DocumentWriter()
        {
            _memoryStream = new MemoryStream();
            _document = WordprocessingDocument.Create(_memoryStream, WordprocessingDocumentType.Document);

            var mainPart = _document.AddMainDocumentPart();

            var body = new Body();
            mainPart.Document = new Document(body);
        }
        /// <summary>
        /// Append a paragraph to the document
        /// </summary>
        /// <param name="sentence"></param>
        public void AddParagraph(string sentence)
        {
            List<Run> runList = ListOfStringToRunList(new List<string> { sentence });
            AddParagraph(runList);
        }
        /// <summary>
        /// Append multiple paragraphs to the document
        /// </summary>
        /// <param name="sentences"></param>
        public void AddParagraph(List<string> sentences)
        {
            List<Run> runList = ListOfStringToRunList(sentences);
            AddParagraph(runList);
        }
        /// <summary>
        /// Append paragraphs from a list of Run objects.
        /// </summary>
        /// <param name="runList"></param>
        public void AddParagraph(List<Run> runList)
        {
            var para = new Paragraph();
            foreach (Run runItem in runList)
            {
                para.AppendChild(runItem);
            }

            var body = _document.MainDocumentPart.Document.Body;
            body.AppendChild(para);
        }
        /// <summary>
        /// Append to the document a list of sentences (list of string) and create bullet list
        /// </summary>
        /// <param name="sentences"></param>
        public void AddBulletList(List<string> sentences)
        {
            var runList = ListOfStringToRunList(sentences);

            AddBulletList(runList);
        }
        /// <summary>
        /// Append to the document a list of sentences (list of Run) and create bullet list
        /// </summary>
        /// <param name="runList"></param>
        public void AddBulletList(List<Run> runList)
        {
            // Introduce bulleted numbering in case it will be needed at some point
            NumberingDefinitionsPart numberingPart = _document.MainDocumentPart.NumberingDefinitionsPart;
            if (numberingPart == null)
            {
                numberingPart = _document.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart001");
                var element = new Numbering();
                element.Save(numberingPart);
            }

            // Insert an AbstractNum into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK productivity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last AbstractNum and BEFORE the first NumberingInstance or we will get a validation error.
            var abstractNumberId = numberingPart.Numbering.Elements<AbstractNum>().Count() + 1;

            var abstractLevel = new Level(new NumberingFormat()
            {
                Val = NumberFormatValues.Bullet
            }, new LevelText() { Val = "·" }) { LevelIndex = 0 };

            var abstractNum1 = new AbstractNum(abstractLevel) { AbstractNumberId = abstractNumberId };

            if (abstractNumberId == 1)
            {
                numberingPart.Numbering.Append(abstractNum1);
            }
            else
            {
                var lastAbstractNum = numberingPart.Numbering.Elements<AbstractNum>().Last();
                numberingPart.Numbering.InsertAfter(abstractNum1, lastAbstractNum);
            }

            // Insert an NumberingInstance into the numbering part numbering list.  The order seems to matter or it will not pass the 
            // Open XML SDK Productity Tools validation test.  AbstractNum comes first and then NumberingInstance and we want to
            // insert this AFTER the last NumberingInstance and AFTER all the AbstractNum entries or we will get a validation error.
            var numberId = numberingPart.Numbering.Elements<NumberingInstance>().Count() + 1;
            var numberingInstance1 = new NumberingInstance() { NumberID = numberId };
            var abstractNumId1 = new AbstractNumId() { Val = abstractNumberId };
            numberingInstance1.Append(abstractNumId1);

            if (numberId == 1)
            {
                numberingPart.Numbering.Append(numberingInstance1);
            }
            else
            {
                var lastNumberingInstance = numberingPart.Numbering.Elements<NumberingInstance>().Last();
                numberingPart.Numbering.InsertAfter(numberingInstance1, lastNumberingInstance);
            }

            Body body = _document.MainDocumentPart.Document.Body;

            foreach (Run runItem in runList)
            {
                // Create items for paragraph properties
                var numberingProperties = new NumberingProperties(new NumberingLevelReference()
                {
                    Val = 0
                }, new NumberingId() { Val = numberId });

                var spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };  // Get rid of space between bullets
                var indentation = new Indentation() { Left = "720", Hanging = "360" };  // correct indentation 

                var paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                var runFonts1 = new RunFonts() { Ascii = "Symbol", HighAnsi = "Symbol" };
                paragraphMarkRunProperties1.Append(runFonts1);

                // create paragraph properties
                var paragraphProperties = new ParagraphProperties(
                    numberingProperties, 
                    spacingBetweenLines1, 
                    indentation, 
                    paragraphMarkRunProperties1);

                // Create paragraph 
                var newPara = new Paragraph(paragraphProperties);

                // Add run to the paragraph
                newPara.AppendChild(runItem);

                // Add one bullet item to the body
                body.AppendChild(newPara);
            }
        }
        public void Dispose()
        {
            CloseAndDisposeOfDocument();

            if (_memoryStream != null)
            {
                _memoryStream.Dispose();
                _memoryStream = null;
            }
        }
        /// <summary>
        /// Save document.
        /// </summary>
        /// <param name="pFileName">Path and file name to save to</param>
        public void SaveToFile(string pFileName)
        {
            if (_document != null)
            {
                CloseAndDisposeOfDocument();
            }

            if (_memoryStream == null)
                throw new ArgumentException("This object has already been disposed of so you cannot save it!");

            using (var fs = File.Create(pFileName))
            {
                _memoryStream.WriteTo(fs);
            }
        }
        /// <summary>
        /// Dispose of document object.
        /// </summary>
        private void CloseAndDisposeOfDocument()
        {
            if (_document != null)
            {
                _document.Close();
                _document.Dispose();
                _document = null;
            }
        }
        private static List<Run> ListOfStringToRunList(List<string> sentences)
        {
            var runList = new List<Run>();
            foreach (var item in sentences)
            {
                var newRun = new Run();
                newRun.AppendChild(new Text(item));
                runList.Add(newRun);
            }

            return runList;
        }
    }
}
