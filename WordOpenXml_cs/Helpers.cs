using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using BlipFill = DocumentFormat.OpenXml.Drawing.Pictures.BlipFill;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using Color = System.Drawing.Color;
using InsideHorizontalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder;
using InsideVerticalBorder = DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using NonVisualDrawingProperties = DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties;
using NonVisualGraphicFrameDrawingProperties = DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties;
using NonVisualPictureDrawingProperties = DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties;
using NonVisualPictureProperties = DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Picture = DocumentFormat.OpenXml.Drawing.Pictures.Picture;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using ShapeProperties = DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties;
using TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;

namespace WordOpenXml_cs
{
    /// <summary>
    /// Code by Karen Payne MVP
    /// </summary>
    public static class Helpers
    {
        public enum PropertyTypes : int
        {
            YesNo,
            Text,
            DateTime,
            NumberInteger,
            NumberDouble
        }

        /// <summary>
        /// Get web color to assign to assign to a Word object color property
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public static string ColorConverter(Color color)
        {
            return color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
        }
        #region Image processing

        /// <summary>
        /// MSDN code sample
        /// https://docs.microsoft.com/en-us/office/open-xml/how-to-insert-a-picture-into-a-word-processing-document
        /// 
        /// There is one issue, EditId attribute is not in the current SDK. This was learned by validating a document
        /// using <see cref="ValidateCorruptedWordDocument"/>. It has been commented out in the event someone wants
        /// to work with this code with an earlier release of the SDK.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="relationshipId"></param>
        /// <param name="pWidth"></param>
        /// <param name="pHeight"></param>
        public static void AddImageToBody(WordprocessingDocument document, string relationshipId, int pWidth, int pHeight)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new Inline(
                         new Extent() { Cx = pWidth, Cy = pHeight },
                         new EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new NonVisualGraphicFrameDrawingProperties(
                             new GraphicFrameLocks()
                             {
                                 NoChangeAspect = true
                             }),
                         new Graphic(
                             new GraphicData(
                                 new Picture(
                                     new NonVisualPictureProperties(
                                         new NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new NonVisualPictureDrawingProperties()),
                                     new BlipFill(
                                         new Blip(
                                             new BlipExtensionList(
                                                 new BlipExtension()
                                                 {
                                                     Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             BlipCompressionValues.Print
                                         },
                                         new Stretch(
                                             new FillRectangle())),
                                     new ShapeProperties(
                                         new Transform2D(
                                             new Offset() { X = 0L, Y = 0L },
                                             new Extents() { Cx = pWidth, Cy = pHeight }),
                                         new PresetGeometry(
                                             new AdjustValueList()
                                         )
                                         { Preset = ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U //,EditId = "50D07946" this is from an older release of the SDK
                     });

            // Append the reference to body, the element should be in a Run.
            document.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        #endregion

        /// <summary>
        /// Create properties for a Word table
        /// </summary>
        /// <returns>TableProperties for setting table borders</returns>
        public static TableProperties CreateTableProperties()
        {
            return new TableProperties(
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 })
            );
        }
        /// <summary>
        /// This comes in handy when learning how to code, should be placed at the
        /// end of a method to validate your code is good.
        /// </summary>
        /// <param name="pFileName"></param>
        /// <returns></returns>
        public static int ValidateWordDocument(string pFileName)
        {
            var errorCount = 0;

            using (var document = WordprocessingDocument.Open(pFileName, true))
            {
                try
                {
                    var validator = new OpenXmlValidator();
                    foreach (ValidationErrorInfo error in validator.Validate(document))
                    {
                        errorCount++;
                        Console.WriteLine("Error " + errorCount);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("Error Type: " + error.ErrorType);
                        Console.WriteLine("Node: " + error.Node);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }

                document.Close();
            }

            return errorCount;
        }
        /// <summary>
        /// Does what it says, corrupt a document :-)
        /// </summary>
        /// <param name="pFileName"></param>
        public static void ValidateCorruptedWordDocument(string pFileName)
        {
            var errorCount = 0;
            // Insert some text into the body, this would cause Schema Error
            using (var document = WordprocessingDocument.Open(pFileName, true))
            {
                // Insert some text into the body, this would cause Schema Error
                var body = document.MainDocumentPart.Document.Body;
                Run run = new Run(new Text("some text"));

                body.Append(run);

                try
                {
                    var validator = new OpenXmlValidator();
                    foreach (ValidationErrorInfo error in validator.Validate(document))
                    {
                        errorCount++;
                        Console.WriteLine("Error " + errorCount);
                        Console.WriteLine("Description: " + error.Description);
                        Console.WriteLine("Error Type: " + error.ErrorType);
                        Console.WriteLine("Node: " + error.Node);
                        Console.WriteLine("Path: " + error.Path.XPath);
                        Console.WriteLine("Part: " + error.Part.Uri);
                        Console.WriteLine("-------------------------------------------");
                    }

                    Console.WriteLine($"count={errorCount}");
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

        }
    }
}
