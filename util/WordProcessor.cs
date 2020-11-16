using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

namespace open_xml_demo.util
{
    class WordProcessor
    {
        private WordprocessingDocument curDoc;
        public WordProcessor(MemoryStream stream)
        {
            curDoc = WordprocessingDocument.Open(stream, true);
        }

        public void Save(string saveLoc)
        {
            curDoc.SaveAs(saveLoc);
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
            curDoc.MainDocumentPart.Document.Save();
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
                    curDoc.Save();
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
