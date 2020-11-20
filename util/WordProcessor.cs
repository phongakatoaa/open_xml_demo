using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System;
using System.IO;

namespace open_xml_demo.util
{
    class WordProcessor
    {
        private WordprocessingDocument curDoc;
        private ImageHelper imageHelper = new ImageHelper(1U);

        public WordProcessor(MemoryStream stream)
        {
            OpenSettings os = new OpenSettings
            {
                AutoSave = false
            };
            curDoc = WordprocessingDocument.Open(stream, true, os);
        }

        public WordProcessor(String fileName)
        {
            OpenSettings os = new OpenSettings
            {
                AutoSave = false
            };
            curDoc = WordprocessingDocument.Open(fileName, true, os);
        }

        public void Save(string saveLoc)
        {
            curDoc.SaveAs(saveLoc);
        }

        public void Close()
        {
            curDoc.Dispose();
        }

        public void FindAndReplace(string tag, string value)
        {
            foreach (var p in curDoc.MainDocumentPart.Document.Descendants<Text>())
            {
                if (p.Text.Contains(tag))
                {
                    p.Text = p.Text.Replace(tag, value);
                }
            }
        }

        public void InsertTable(string placeholder)
        {
            foreach (var p in curDoc.MainDocumentPart.Document.Descendants<Paragraph>())
            {
                if (p.InnerText.Contains(placeholder))
                {
                    Table table = new Table();
                    TableProperties props = new TableProperties(
                        new TableBorders(
                            new TopBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 12
                            },
                            new BottomBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 12
                            },
                            new LeftBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 12
                            },
                            new RightBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 12
                            },
                            new InsideHorizontalBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 12
                            },
                            new InsideVerticalBorder
                            {
                                Val = new EnumValue<BorderValues>(BorderValues.Single),
                                Size = 12
                            }
                        ),
                        new TableWidth()
                        {
                            Width = "5000",
                            Type = TableWidthUnitValues.Pct
                        }
                    );

                    table.AppendChild<TableProperties>(props);

                    var row_1 = new TableRow();
                    var cell_1_1 = InsertCell(row_1, "abc");
                    var cell_1_2 = InsertCell(row_1, "def");
                    var cell_1_3 = InsertCell(row_1, "ghi");

                    var row_2 = new TableRow();

                    var cell_2_1 = InsertCell(row_2, "jkl");
                    var cell_prop_2_1 = new TableCellProperties();
                    cell_prop_2_1.Append(new HorizontalMerge()
                    {
                        Val = MergedCellValues.Restart
                    });
                    cell_2_1.Append(cell_prop_2_1);

                    var cell_2_2 = InsertCell(row_2, "mno");
                    var cell_prop_2_2 = new TableCellProperties();
                    cell_prop_2_2.Append(new HorizontalMerge()
                    {
                        Val = MergedCellValues.Continue
                    });
                    cell_2_2.Append(cell_prop_2_2);

                    var cell_2_3 = InsertCell(row_2, "pqr");

                    table.Append(row_1, row_2);
                    curDoc.MainDocumentPart.Document.Body.ReplaceChild(table, p);
                }
            }
        }

        public void InsertImage(string placeholder, string fileName, int width, int height)
        {
            ImagePart imagePart = curDoc.MainDocumentPart.AddImagePart(ImagePartType.Png);
            using (FileStream fileStream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(fileStream);
            }
            string relationshipId = curDoc.MainDocumentPart.GetIdOfPart(imagePart);
            foreach (var p in curDoc.MainDocumentPart.Document.Body.Descendants<Paragraph>())
            {
                Console.WriteLine(p.InnerText);
                if (p.InnerText.Equals(placeholder))
                {
                    var element = imageHelper.GetImageElement(relationshipId, fileName, Guid.NewGuid().ToString(), width, height);
                    curDoc.MainDocumentPart.Document.Body.ReplaceChild(new Paragraph(new Run(element)), p);
                }
            }

            foreach (var t in curDoc.MainDocumentPart.Document.Body.Descendants<Table>())
            {
                foreach (var r in t.Descendants<TableRow>())
                {
                    foreach(var c in r.Descendants<TableCell>())
                    {
                        foreach(var p in c.Descendants<Paragraph>())
                        {
                            Console.WriteLine(p.InnerText);
                            if (p.InnerText.Equals(placeholder))
                            {
                                var element = imageHelper.GetImageElement(relationshipId, fileName, Guid.NewGuid().ToString(), width, height);
                                c.ReplaceChild(new Paragraph(new Run(element)), p);
                            }
                        }
                    }
                }
            }
        }

        private TableCell InsertCell(TableRow row, string cellValue)
        {
            TableCell cell = new TableCell();
            cell.Append(new Paragraph(new Run(new Text(cellValue))));
            row.Append(cell);
            return cell;
        }
    }

}
