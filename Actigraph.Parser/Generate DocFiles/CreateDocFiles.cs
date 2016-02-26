using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Xsl;
using Microsoft.Office.Interop.Word;
using System = Microsoft.Office.Interop.Word.System;

namespace Actigraph.Parser.Generate_DocFiles
{
    public class CreateDocFiles
    {

        public string fileExtension;
        private List<TableData> tableData1 = new List<TableData>();
        private List<TableData> tableData2 = new List<TableData>();
        private SummaryTableData[] tableData3 = null;
        private SubjectValidAverages subjectValidAverages;
        private long ThresholdCutOffValid = 0L;
        public CreateDocFiles(long thresholdCutOff)
        {
            ThresholdCutOffValid = thresholdCutOff;
        }
        public void CreateReports(SubjectRecords subjectRecords)
        {
            try
            {
                subjectValidAverages = subjectRecords.SubjectValidAverages;
                GetTable1Data(subjectRecords.SubjectData);
                GetTable2Data(subjectRecords);
                tableData3 = GetTable3Data(subjectRecords);
                Create();
            }
            catch (Exception exp)
            {
                throw exp;
            }

        }

        private SummaryTableData[] GetTable3Data(SubjectRecords subjectRecords)
        {
            return (from rec in subjectRecords.SubjectData
                select new SummaryTableData()
                {
                    Date = rec.Date.ToShortDateString(),
                    Day_of_Week = rec.DayofWeek,
                    Wear_Time = Math.Round(rec.Time),
                    Movements_Per_Minute = Math.Round(rec.VectorMagnitudeCpm),
                    Steps = rec.StepsCount,
                    Sedentary = Math.Round(rec.Sedentary),
                    Light = Math.Round(rec.Light),
                    LifeStyle = Math.Round(rec.Lifestyle),
                    Moderate = Math.Round(rec.Moderate)
                }).OrderBy(p => p.Date).ToArray();
        }



        private void GetTable1Data(IEnumerable<SubjectData> subjectData)
        {
            var data = (from rec in subjectData
                select new
                {
                    rec.ID,
                    rec.Cottage,
                    rec.Age,
                    rec.Gender
                }).FirstOrDefault();

            foreach (var item in (data.GetType().GetProperties()))
            {
                tableData1.Add(new TableData()
                {
                    Name = item.Name,
                    Value = Convert.ToString(item.GetValue(data))
                });
            }
            //return tableData;
        }

        private void GetTable2Data(SubjectRecords subjectRecords)
        {
            var dateRange = subjectRecords.SubjectData.OrderBy(p => p.Date).FirstOrDefault().Date.ToShortDateString()
                            + " to " +
                            subjectRecords.SubjectData.OrderByDescending(p => p.Date)
                                .FirstOrDefault()
                                .Date.ToShortDateString();
            var validDays = subjectRecords.SubjectValidData.Count();
            var invalidDays = subjectRecords.SubjectData.Count() - validDays;
            var avgAllDaysWearTime = subjectRecords.SubjectData.Select(p => p.Time).Average();
            var avgValidaysWearTime = subjectRecords.SubjectValidAverages.AvgTime;
            var avgMovements = subjectRecords.SubjectValidData.Select(p => p.VectorMagnitudeCpm).Average();
            var avgSteps = subjectRecords.SubjectValidData.Select(p => p.StepsCount).Average();

            var data = new
            {
                Date_Range = dateRange,
                Valid_Days = validDays,
                Invalid_Days = invalidDays,
                Average_Wear_Time_All_Days = Math.Round(avgAllDaysWearTime),
                Average_Wear_Time_Valid_Days = Math.Round(avgValidaysWearTime),
                Average_Movements_Valid_Days = Math.Round(avgMovements),
                Average_Steps_Valid_Days = Math.Round(avgSteps)
            };

            foreach (var item in (data.GetType().GetProperties()))
            {
                tableData2.Add(new TableData()
                {
                    Name = item.Name.Replace('_', ' '),
                    Value = Convert.ToString(item.GetValue(data))
                });
            }
        }

        private void Create()
        {
            object missing = Missing.Value;
            Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
            var document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            try
            {
                //Create an instance for word app

                //Set animation status for word application
                // winword.ShowAnimation = false;

                //Set status for word application is to be visible or not.
                winword.Visible = false;

                //Create a missing variable for missing value

                //Create a new document
                document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                document.PageSetup.TopMargin = 35f;
                document.PageSetup.BottomMargin = 35f;
                document.PageSetup.LeftMargin = 35f;
                document.PageSetup.RightMargin = 35f;
                document.ShowGrammaticalErrors = false;
                document.GrammarChecked = false;
                document.SpellingChecked = false;
                //Add header into the document
                foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                {
                    //Get the header range and add the header details.
                    var headerRange =
                        section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment =
                        Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                    headerRange.Font.Size = 20;
                    headerRange.Bold = 2;
                    headerRange.Text = "Actigraph Weekly Summary Report" + Environment.NewLine;
                }

                document.Content.SetRange(0, 0);
                document.Content.Text = "Report Date: " + DateTime.Now.ToShortDateString() + Environment.NewLine;
                document.Content.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                document.Content.Font.Size = 10;
                document.Content.Bold = 1;
                document.Content.Font.Name = "Microsoft YaHei UI";
                //Add paragraph with Heading 1 style
                var para1 = document.Content.Paragraphs.Add(ref missing);
                object styleStrong = "Strong";
                para1.Range.set_Style(ref styleStrong);
                para1.Range.Text = "Subject Details:";
                para1.Range.Font.Color = WdColor.wdColorDarkBlue;
                para1.Range.Underline = WdUnderline.wdUnderlineSingle;
                para1.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para1.Range.Font.Size = 14;
                para1.Range.InsertParagraphAfter();


                var para2 = document.Content.Paragraphs.Add(ref missing);
                para2.Range.InsertParagraphAfter();


                var firstTable = document.Tables.Add(para2.Range, tableData1.Count, 2, ref missing, ref missing);

                CreateTable(ref firstTable, tableData1, 150f);
                firstTable.Range.Paragraphs.SpaceAfter = 1;
                //Table 2
                var para3 = document.Content.Paragraphs.Add(ref missing);

                para3.Range.set_Style(ref styleStrong);
                para3.Range.Text = "Subject Actigraph Details (Weekly):";
                para3.Range.Underline = WdUnderline.wdUnderlineSingle;
                para3.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para3.Range.Font.Size = 14;
                para3.Range.Font.Color = WdColor.wdColorDarkBlue;
                para3.Range.InsertParagraphAfter();

                var avgSubjectTable = document.Tables.Add(para3.Range, tableData2.Count, 2, ref missing, ref missing);
                CreateTable(ref avgSubjectTable, tableData2, 285f);
                avgSubjectTable.Range.Paragraphs.SpaceAfter = 1;
                document.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                //Table3
                var para4 = document.Content.Paragraphs.Add(ref missing);

                para4.Range.set_Style(ref styleStrong);
                para4.Range.Text = "Subject Summary Details (Weekly):";
                para3.Range.Font.Color = WdColor.wdColorDarkBlue;
                para4.Range.Underline = WdUnderline.wdUnderlineSingle;
                para4.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                para4.Range.Font.Size = 14;
                para4.Range.InsertParagraphAfter();

                var subjectSummaryTable = document.Tables.Add(para4.Range, tableData3.Count()+1, 9, ref missing,
                    ref missing);
                subjectSummaryTable.Borders.Enable = 1;
                CreateSummaryTable(ref subjectSummaryTable, tableData3, 150f);
                CreateSummaryLastRow(ref subjectSummaryTable);
                subjectSummaryTable.Range.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
                subjectSummaryTable.Range.Paragraphs.SpaceAfter = 1;
                CreateDocument(document);
               
                //  MessageBox.Show("Document created successfully !");


            }
            catch (Exception exp)
            {
                document.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges,
                    Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat,
                    false);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                throw exp;
            }
            finally
            {
                document.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges,
                    Microsoft.Office.Interop.Word.WdOriginalFormat.wdOriginalDocumentFormat,
                    false);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
            }
        }

        private void CreateDocument(Document document)
        {
           
            Object oMissing = Missing.Value;
            var subjectId = tableData1.Where(i => i.Name == "ID").ToArray()[0].Value;
                var folderName = DirectoryStructure.CreateSubjectFolder(subjectId);
                object fileName = Path.Combine(folderName, subjectId + "-" + DateTime.Now.ToShortDateString() + fileExtension);
                if (File.Exists(fileName.ToString()))
                {
                fileName = Path.Combine(folderName,
                        Guid.NewGuid() +
                        Path.GetExtension(fileName.ToString()));
                }
                if(fileExtension.ToUpper() == ".PDF")
                document.ExportAsFixedFormat(fileName.ToString(), WdExportFormat.wdExportFormatPDF, false, WdExportOptimizeFor.wdExportOptimizeForOnScreen,
                WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, true, true,
                WdExportCreateBookmarks.wdExportCreateHeadingBookmarks, true, true, false, ref oMissing);
                else
                {
                    document.SaveAs(fileName);
                }
            }

        private void CreateTable(ref Table firstTable, List<TableData> tableData, float width)
        {
            foreach (Row row in firstTable.Rows)
            {
                
                foreach (Cell cell in row.Cells)
                {
                    cell.Height = 7;
                    cell.Range.Font.Size = 10;
                    cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                    cell.Range.Underline = WdUnderline.wdUnderlineNone;
                    cell.LeftPadding = 5;
                    cell.RightPadding = 5;
                    cell.Range.Font.Color=WdColor.wdColorBlack;

                    if (cell.ColumnIndex == 1)
                    {
                        cell.LeftPadding = 50;
                        cell.FitText = false;
                        cell.Width = width;
                        cell.Range.Font.Bold = 1;
                        cell.Range.Text = tableData[row.Index - 1].Name;
                    }
                    else
                    {
                        cell.Range.Font.Bold = 0;
                        cell.Range.Text = tableData[row.Index - 1].Value;
                    }
                }
            }
        }

        private void CreateSummaryLastRow(ref Table subjectSummaryTable)
        {
            Object missingObj = Missing.Value;
            subjectSummaryTable.Rows.Add(missingObj);
            subjectSummaryTable.Rows[subjectSummaryTable.Rows.Count].Shading.BackgroundPatternColor=WdColor.wdColorWhite;
            foreach (Cell cell in subjectSummaryTable.Rows[subjectSummaryTable.Rows.Count].Cells)
            {
                cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                cell.Width = 90;
                cell.Range.Font.Size = 10;
                cell.Range.Bold = 0;
                cell.Range.Underline = WdUnderline.wdUnderlineNone;
                cell.Range.Font.Color = WdColor.wdColorBlack;
                switch (cell.ColumnIndex)
                {
                    case 1:
                        cell.Range.Text = "Avg for valid days(" + subjectValidAverages.ValidDays + ")"; 
                        break;
                    case 2:
                        cell.Range.Text = "N/A";
                        break;
                    case 3:
                        cell.Range.Text =
                            tableData2.Where(t=>t.Name== "Average Wear Time Valid Days").ToArray()[0].Value;
                        break;
                    case 4:
                        cell.Range.Text =
                            tableData2.Where(t => t.Name == "Average Movements Valid Days").ToArray()[0].Value;
                        break;
                    case 5:
                        cell.Range.Text = tableData2.Where(t => t.Name == "Average Steps Valid Days").ToArray()[0].Value; 
                        break;
                    case 6:
                        cell.Range.Text =
                            subjectValidAverages.AvgSedentary.ToString();
                        break;
                    case 7:
                        cell.Range.Text = subjectValidAverages.AvgLight.ToString();
                        break;
                    case 8:
                        cell.Range.Text =
                            subjectValidAverages.AvgLifestyle.ToString();
                        break;
                    case 9:
                        cell.Range.Text =
                            subjectValidAverages.AvgModerate.ToString();
                        break;
                    default:
                        break;
                }

            }
            }
        
        private void CreateSummaryTable(ref Table subjectSummaryTable, SummaryTableData[] tableData, float width)
        {
            foreach (Row row in subjectSummaryTable.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    //Header row
                    if (cell.RowIndex == 1)
                    {
                        switch (cell.ColumnIndex)
                        {
                            case 1:
                                cell.Range.Text = "Date";
                                // cell.Range.Text = tableData[row.Index - 2].Date;
                                break;
                            case 2:
                                cell.Range.Text = "Day of Week";
                                break;
                            case 3:
                                cell.Range.Text = "Wear Time";
                                break;
                            case 4:
                                cell.Range.Text = "Movements Per Minute";
                                break;
                            case 5:
                                cell.Range.Text = "Steps";
                                break;
                            case 6:
                                cell.Range.Text = "Sedentary";
                                break;
                            case 7:
                                cell.Range.Text = "Light";
                                break;
                            case 8:
                                cell.Range.Text = "LifeStyle";
                                break;
                            case 9:
                                cell.Range.Text = "Moderate";
                                break;
                            default:
                                break;
                        }

                        cell.Range.Font.Bold = 1;
                        //other format properties goes heres
                        cell.Range.Font.Size = 10;
                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                            
                        cell.Shading.BackgroundPatternColor = WdColor.wdColorLightBlue;
                        cell.Range.Font.Color = WdColor.wdColorWhite;
                        //Center alignment for the Header cells
                        cell.Range.Underline = WdUnderline.wdUnderlineNone;
                        cell.Width = 90;
                        cell.Height = 7;
                    }
                    //Data row
                    else
                    {
                        cell.Width = 90;
                        cell.Range.Bold = 0;
                        cell.Range.Underline = WdUnderline.wdUnderlineNone;
                        cell.Range.Font.Size = 10;
                        cell.Range.Font.Color = WdColor.wdColorBlack;
                        switch (cell.ColumnIndex)
                        {
                            case 1:
                                cell.Range.Text = tableData[row.Index - 2].Date;
                                break;
                            case 2:
                                cell.Range.Text = tableData[row.Index - 2].Day_of_Week;
                                break;
                            case 3:
                                cell.Range.Text =
                                    tableData[row.Index - 2].Wear_Time.ToString(CultureInfo.InvariantCulture);
                                if (tableData[row.Index - 2].Wear_Time < ThresholdCutOffValid)
                                {
                                    Color c = Color.LightGray;
                                    var wdc = (WdColor)(c.R + 0x100 * c.G + 0x10000 * c.B);
                                    row.Shading.BackgroundPatternColor = wdc;
                                }
                                break;
                            case 4:
                                cell.Range.Text =
                                    tableData[row.Index - 2].Movements_Per_Minute.ToString(CultureInfo.InvariantCulture);
                                break;
                            case 5:
                                cell.Range.Text = tableData[row.Index - 2].Steps.ToString();
                                break;
                            case 6:
                                cell.Range.Text =
                                    tableData[row.Index - 2].Sedentary.ToString(CultureInfo.InvariantCulture);
                                break;
                            case 7:
                                cell.Range.Text = tableData[row.Index - 2].Light.ToString(CultureInfo.InvariantCulture);
                                break;
                            case 8:
                                cell.Range.Text =
                                    tableData[row.Index - 2].LifeStyle.ToString(CultureInfo.InvariantCulture);
                                break;
                            case 9:
                                cell.Range.Text =
                                    tableData[row.Index - 2].Moderate.ToString(CultureInfo.InvariantCulture);
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
          //  subjectSummaryTable.Rows.Add(ref missingObj);

        }
    }
}
