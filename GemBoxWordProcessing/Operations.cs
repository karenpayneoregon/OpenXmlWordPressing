using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using GemBox.Document;
using GemBox.Document.Tables;

namespace GemBoxWordProcessing
{
    /// <summary>
    /// https://www.gemboxsoftware.com/
    ///
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
            ComponentInfo.SetLicense("FREE-LIMITED-KEY");
            ComponentInfo.FreeLimitReached += (sender, e) => e.FreeLimitReachedAction = FreeLimitReachedAction.ContinueAsTrial;
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

        public bool CreateEmptyDocument(string pFileName)
        {
            mHasException = false;

            var fileName = Path.Combine(DocumentFolder, pFileName);
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            try
            {
                var document = new DocumentModel();
                document.Save(fileName);
            }
            catch (Exception e)
            {
                mHasException = true;
                mLastException = e;
            }

            return IsSuccessFul;
        }
        /// <summary>
        /// Create new document with a simple paragraph.
        /// </summary>
        /// <param name="pFileName"></param>
        /// <returns></returns>
        public bool CreateDocumentWithSimpleParagraph(string pFileName)
        {
            mHasException = false;

            var fileName = Path.Combine(DocumentFolder, pFileName);
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            var paragraphText =
                "The most basic unit of block-level content within a Word processing document, paragraphs are " +
                "stored using the <p> element. A paragraph defines a distinct division of content that begins on " +
                "a new line. A paragraph can contain three pieces of information: optional paragraph properties, " +
                "inline content (typically runs), and a set of optional revision IDs used to compare the content " +
                "of two documents.";

            try
            {
                var document = new DocumentModel();
                document.DefaultCharacterFormat.FontColor = Color.Brown;
                document.Sections.Add(new Section(document,new Paragraph(document, paragraphText)));

                document.Save(fileName);
            }
            catch (Exception e)
            {
                mHasException = true;
                mLastException = e;
            }

            return IsSuccessFul;
        }
        /// <summary>
        /// Create a new document with unordered list.
        /// </summary>
        /// <param name="pFileName"></param>
        /// <returns></returns>
        public bool CreateDocumentWithUnoderedList(string pFileName)
        {
            mHasException = false;

            var fileName = Path.Combine(DocumentFolder, pFileName);
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            try
            {
                var document = new DocumentModel();

                var section = new Section(document, new Paragraph(document, "This is a spacing paragraph 1."));
                document.Sections.Add(section);

                var bulletList = new ListStyle(ListTemplateType.Bullet);

                var fruitList = new List<string>() { "Apple", "Banana", "Carrot" };

                foreach (var fruit in fruitList)
                {
                    section.Blocks.Add(
                        new Paragraph(document, fruit)
                        {
                            ParagraphFormat = new ParagraphFormat() { NoSpaceBetweenParagraphsOfSameStyle = true },
                            ListFormat = new ListFormat() { Style = bulletList }
                        });
                }

                document.Sections.Add(new Section(document, new Paragraph(document, "This is a spacing paragraph 2.")));

                var animalList = new List<string>() { "Dog", "Cat", "Bear" };
                foreach (var animal in animalList)
                {
                    section.Blocks.Add(
                        new Paragraph(document, animal)
                        {
                            ParagraphFormat = new ParagraphFormat() { NoSpaceBetweenParagraphsOfSameStyle = true },
                            ListFormat = new ListFormat() { Style = bulletList }
                        });
                }

                document.Sections.Add(new Section(document, new Paragraph(document, "Gembox Document is simple to use")));

                document.Save(fileName);
            }
            catch (Exception e)
            {
                mHasException = true;
                mLastException = e;
            }


            return IsSuccessFul;
        }
        /// <summary>
        /// Create new document with an image.
        /// </summary>
        /// <param name="pFileName"></param>
        /// <returns></returns>
        public bool CreateDocumentWithMultipleParagraphAndImage(string pFileName)
        {
            mHasException = false;

            var fileName = Path.Combine(DocumentFolder, pFileName);
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            var paragraphText =
                "The most basic unit of block-level content within a Word processing document, paragraphs are " +
                "stored using the <p> element. A paragraph defines a distinct division of content that begins on " +
                "a new line. A paragraph can contain three pieces of information: optional paragraph properties, " +
                "inline content (typically runs), and a set of optional revision IDs used to compare the content " +
                "of two documents.";

            try
            {

                var document = new DocumentModel();
                var imageFileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "F2.png");

                // Built-in styles can be created using Style.CreateStyle() method.
                var title = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Title, document);


                var emphasis = new CharacterStyle("Emphasis");
                emphasis.CharacterFormat.Italic = true;

                // First add style to the document, then use it.
                document.Styles.Add(title);
                document.Styles.Add(emphasis);

                document.Sections.Add(
                    new Section(document,
                        new Paragraph(document, "Image example")
                        {
                            ParagraphFormat = new ParagraphFormat()
                            {
                                Style = title
                            }
                        },
                        new Paragraph(document,
                            new Run(document, paragraphText)
                            {
                                CharacterFormat = new CharacterFormat()
                                {
                                    Style = (CharacterStyle)document.Styles.GetOrAdd(StyleTemplateType.Strong)
                                }
                            },
                            new Run(document, " But let's check out the car.")
                            {
                                CharacterFormat = new CharacterFormat()
                                {
                                    Style = emphasis
                                }
                            })
                        ));


                var section = new Section(document);
                document.Sections.Add(section);

                var paragraph = new Paragraph(document);
                section.Blocks.Add(paragraph);

                var miataPicture = new Picture(document, imageFileName, 579, 300, LengthUnit.Pixel);
                paragraph.Inlines.Add(miataPicture);

                document.Sections.Add(new Section(document, new Paragraph(document, "We just inserted an image.")));

                document.Save(fileName);


            }
            catch (Exception e)
            {
                mHasException = true;
                mLastException = e;
            }

            return IsSuccessFul;
        }
        /// <summary>
        /// Create new document with a table
        /// </summary>
        /// <param name="pFileName"></param>
        /// <returns></returns>
        public bool CreateDocumentWithTableSimple(string pFileName)
        {

            mHasException = false;

            var fileName = Path.Combine(DocumentFolder, pFileName);
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }

            var document = new DocumentModel();

            var tableRowCount = 10;
            var tableColumnCount = 5;

            var table = new Table(document)
            {
                TableFormat =
                {
                    PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage)
                }
            };

            for (var i = 0; i < tableRowCount; i++)
            {
                var row = new TableRow(document);
                table.Rows.Add(row);

                for (var j = 0; j < tableColumnCount; j++)
                {
                    var para = new Paragraph(document, $"Cell {i + 1}-{j + 1}");

                    row.Cells.Add(new TableCell(document, para));
                }
            }

            document.Sections.Add(new Section(document, table));

            document.Save(fileName);

            return IsSuccessFul;
        }
    }
}
