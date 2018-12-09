using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.Wordprocessing;
using BottomBorder = DocumentFormat.OpenXml.Drawing.BottomBorder;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Drawing.ParagraphProperties;
using Path = System.IO.Path;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace WordOpenXml_cs
{
    /// <summary>
    /// Code by Karen Payne MVP
    /// </summary>
    public class Operations : BaseExceptionProperties
    {
        /// <summary>
        /// Location where Word documents are created.
        /// </summary>
        /// <remarks>
        /// - First time a build is performed the Documents folder is created by a Post Build event.
        /// - Next time don't create if exists.
        /// </remarks>
        public readonly string DocumentFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents");
        /// <summary>
        /// Remove all documents in base document folder.
        /// </summary>
        /// <remarks>
        /// If there is an exception thrown it's picked up in the class constructor below.
        /// </remarks>
        private void Prepare()
        {
            Directory.GetFiles(DocumentFolder, "*.docx").ToList().ForEach(File.Delete);
        }
        /// <summary>
        /// Create instance of class with option to remove files from base document folder.
        /// </summary>
        /// <param name="pPurge"></param>
        public Operations(bool pPurge = false)
        {
            try
            {
                if (pPurge)
                {
                    Prepare();
                }
            }
            catch (Exception)
            {
                mHasException = true;
            }
        }
        /// <summary>
        /// Create new empty Word document.
        /// </summary>
        /// <returns>Is valid document</returns>
        public bool CreateEmptyDocument(string pFileName)
        {
            mHasException = false;

            var fileName = Path.Combine(DocumentFolder, pFileName);

            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            using (var document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document();
                mainPart.Document.AppendChild(new Body());

                
                mainPart.Document.Save();
            }

            return Helpers.ValidateWordDocument(fileName) == 0;
        }
        /// <summary>
        /// Create a new document, append a single paragraph, set font and color.
        /// </summary>
        /// <returns>Is valid document</returns>
        public bool CreateDocumentWithSimpleParagraph(string pFileName)
        {
            mHasException = false;

            var fileName = Path.Combine(DocumentFolder, pFileName);
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            using (var document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = document.AddMainDocumentPart();

                mainPart.Document = new Document();

                var body = mainPart.Document.AppendChild(new Body());

                var para = body.AppendChild(new Paragraph());
                Run runPara = para.AppendChild(new Run());

                // Set the font to Arial to the first Run.
                var runProperties = new RunProperties(
                    new RunFonts()
                    {
                        Ascii = "Arial"
                    });

                var color = new Color {Val = Helpers.ColorConverter(System.Drawing.Color.SandyBrown)};
                runProperties.Append(color);
                
                Run run = document.MainDocumentPart.Document.Descendants<Run>().First();
                run.PrependChild<RunProperties>(runProperties);

                var paragraphText = 
                    "The most basic unit of block-level content within a Word processing document, paragraphs are " +
                    "stored using the <p> element. A paragraph defines a distinct division of content that begins on " +
                    "a new line. A paragraph can contain three pieces of information: optional paragraph properties, " + 
                    "inline content (typically runs), and a set of optional revision IDs used to compare the content " +
                    "of two documents.";

                runPara.AppendChild(new Text(paragraphText));


                mainPart.Document.Save();
            }


            return Helpers.ValidateWordDocument(fileName) == 0;

        }

        private void NextLevel_1(string pFileName)
        {
            var fileName = Path.Combine(DocumentFolder, pFileName);
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            using (var document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = document.AddMainDocumentPart();

                mainPart.Document = new Document();

                var body = mainPart.Document.AppendChild(new Body());

                var para = body.AppendChild(new Paragraph());
                Run runPara = para.AppendChild(new Run());

                // Set the font to Arial to the first Run.
                var runProperties = new RunProperties(
                    new RunFonts()
                    {
                        Ascii = "Arial"
                    });

                var color = new Color { Val = Helpers.ColorConverter(System.Drawing.Color.SandyBrown) };
                runProperties.Append(color);

                Run run = document.MainDocumentPart.Document.Descendants<Run>().First();
                run.PrependChild<RunProperties>(runProperties);

                var paragraphText = "Styling paragraph with font color";

                runPara.AppendChild(new Text(paragraphText));


                mainPart.Document.Save();
            }



            Console.WriteLine(Helpers.ValidateWordDocument(fileName));

        }

        /// <summary>
        /// Creates paragraphs followed by unordered list
        /// </summary>
        /// <returns></returns>
        public bool CreateDocumentWithUnoderedList(string pFileName)
        {
            mHasException = false;

            var fileName = Path.Combine(DocumentFolder, pFileName);

            if (File.Exists(fileName))
                File.Delete(fileName);

            var writer = new DocumentWriter();
            writer.AddParagraph("This is a spacing paragraph 1.");
            var fruitList = new List<string>() { "Apple", "Banana", "Carrot" };
            writer.AddBulletList(fruitList);
            writer.AddParagraph("This is a spacing paragraph 2.");

            var animalList = new List<string>() { "Dog", "Cat", "Bear" };
            writer.AddBulletList(animalList);
            writer.AddParagraph("This is a spacing paragraph 3.");

            var stuffList = new List<string>() { "Ball", "Wallet", "Phone" };
            writer.AddBulletList(stuffList);
            writer.AddParagraph("Done.");

            writer.SaveToFile(fileName);

            return Helpers.ValidateWordDocument(fileName) == 0;

        }
      
        /// <summary>
        /// Create a new document, append a single paragraph, set font and color.
        /// - add first paragraph with normal style, color brown
        /// - add second paragraph with normal style, default color
        /// - add a third paragraph with part bold and highlight/shade, part normal.
        /// </summary>
        /// <returns>Is valid document</returns>
        public bool CreateDocumentWithMultipleParagraphAndImage(string pFileName)
        {
            mHasException = false;

            var fileName = Path.Combine(DocumentFolder, pFileName);

            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            using (var document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = document.AddMainDocumentPart();

                mainPart.Document = new Document();

                var body = mainPart.Document.AppendChild(new Body());

                var para = body.AppendChild(new Paragraph());
                Run runPara = para.AppendChild(new Run());

                // Set the font to Arial to the first Run.
                var runProperties = new RunProperties(
                    new RunFonts()
                    {
                        Ascii = "Arial"
                    });

                var color = new Color { Val = Helpers.ColorConverter(System.Drawing.Color.Brown) };
                runProperties.Append(color);

                Run run = document.MainDocumentPart.Document.Descendants<Run>().First();
                run.PrependChild<RunProperties>(runProperties);

                var paragraphText =
                    "The most basic unit of block-level content within a Word processing document, paragraphs are " +
                    "stored using the <p> element. A paragraph defines a distinct division of content that begins on " +
                    "a new line. A paragraph can contain three pieces of information: optional paragraph properties, " +
                    "inline content (typically runs), and a set of optional revision IDs used to compare the content " +
                    "of two documents.";

                runPara.AppendChild(new Text(paragraphText));

                // add second paragraph
                para = body.AppendChild(new Paragraph());

                runPara = para.AppendChild(new Run());

                paragraphText =
                    "A paragraph's properties are specified via the <pPr>element. Some examples of paragraph properties " + 
                    "are alignment, border, hyphenation override, indentation, line spacing, shading, text direction, " + 
                    "and widow/orphan control.";

                runPara.AppendChild(new Text(paragraphText));

                
                // Highlight and bold some text.
                para = body.AppendChild(new Paragraph());
                run = para.AppendChild(new Run());
                runProperties = run.AppendChild(new RunProperties());

                runProperties.AppendChild(new Bold { Val = OnOffValue.FromBoolean(true) });

                var shading = new Shading()
                {
                    Val = ShadingPatternValues.Clear,
                    Fill = Helpers.ColorConverter(System.Drawing.Color.Yellow)
                };

                runProperties.Append(shading);

                run.AppendChild(new Text("This is highlight/bold"));

                // back to normal text
                run = para.AppendChild(new Run());
                run.AppendChild(new Text(", and this text is normal."));

                int imageWidth = 0;
                int imageHeight = 0;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                var imageFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "F2.png");

                using (var image = new System.Drawing.Bitmap(imageFileName))
                {
                    imageWidth = image.Width;
                    imageHeight = image.Height;
                }

                using (var stream = new FileStream(imageFileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                Helpers.AddImageToBody(document, mainPart.GetIdOfPart(imagePart), imageWidth.PixelToEmu(), imageHeight.PixelToEmu());


                para = body.AppendChild(new Paragraph());
                run = para.AppendChild(new Run());
                run.AppendChild(new Text("We just inserted an image."));

                mainPart.Document.Save();

            }

            return Helpers.ValidateWordDocument(fileName) == 0;

        }
        /// <summary>
        /// Create new Word document with a two column table from a fixed array.
        /// </summary>
        /// <returns>Is valid document</returns>
        public bool CreateDocumentWithTableFromArray(string pFileName)
        {

            mHasException = false;

            var names = new[,]
            {
                { "Karen", "Payne" },
                { "Bill", "Smith" },
                { "Jane", "Lebow" },
                { "Jess", "Gallagher" }
            };

            var fileName = Path.Combine(DocumentFolder, pFileName);

            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            using (var document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document();
                mainPart.Document.AppendChild(new Body());

                var table = new Table();

                // set borders
                TableProperties props = Helpers.CreateTableProperties();
                table.AppendChild(props);

                // make table full width of page
                var tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
                table.AppendChild(tableWidth);

                for (var outerIndex = 0; outerIndex <= names.GetUpperBound(0); outerIndex++)
                {
                    var tableRow = new TableRow();
                    for (var innerIndex = 0; innerIndex <= names.GetUpperBound(1); innerIndex++)
                    {
                        var tableColumn = new TableCell();

                        tableColumn.Append(new Paragraph(new Run(new Text(names[outerIndex, innerIndex]))));

                        // Assume you want columns that are automatically sized.
                        tableColumn.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                        tableRow.Append(tableColumn);
                    }
                    table.Append(tableRow);
                }

                mainPart.Document.Body.Append(table);
                mainPart.Document.Save();

            }

            return true;

        }
        public void  ChangeLastNameForSecondRow(string pFileName, string pTextValue)
        {
            var fileName = Path.Combine(DocumentFolder, pFileName);
            if (!File.Exists(fileName))
            {
                CreateDocumentWithTableFromArray(fileName);
            }

            // Open our document, second parameter indicates edit more
            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                // Get the sole table in the document
                Table table = document.MainDocumentPart.Document.Body.Elements<Table>().First();

                // Find the second row in the table.
                TableRow row = table.Elements<TableRow>().ElementAt(1);

                // Find the second cell in the row.
                TableCell cell = row.Elements<TableCell>().ElementAt(1);

                // Find the first paragraph in the table cell.
                Paragraph p = cell.Elements<Paragraph>().First();

                // Find the first run in the paragraph.
                Run run = p.Elements<Run>().First();

                // Set the text for the run.
                Text text = run.Elements<Text>().First();
                text.Text = pTextValue;

                document.Save();
            }
        }
    }
}
