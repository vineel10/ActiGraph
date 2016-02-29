using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Font = iTextSharp.text.Font;
using Rectangle = iTextSharp.text.Rectangle;
using System.Windows.Forms.DataVisualization.Charting;
using Image = iTextSharp.text.Image;

namespace Actigraph.Parser.Generate_DocFiles
{
    public class CreatePdfReports
    {
        public string fileExtension;
        private List<TableData> tableData1 = new List<TableData>();
        private List<TableData> tableData2 = new List<TableData>();
        private SummaryTableData[] tableData3 = null;
        private SubjectRecords SubjectRecords;
        private long ThresholdCutOffValid = 0L;

        public CreatePdfReports(long thresholdCutOff)
        {
            ThresholdCutOffValid = thresholdCutOff;
        }

        public void CreateReports(SubjectRecords subjectRecords)
        {
            try
            {
                SubjectRecords = subjectRecords;
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
                    Moderate = Math.Round(rec.Moderate),
                    Moderate10 = Math.Round(rec.Moderate10),
                }).OrderBy(p => p.Date).ToArray();
        }

        private MemoryStream CreateGraph(SeriesChartType chartType)
        {
            Chart chart1 = new Chart()
            {
                Width = chartType != SeriesChartType.Pie ? 250 : 330,
                Height = chartType == SeriesChartType.Pie ? 280 : 220
            };

            ChartArea chartArea = new ChartArea("chartArea1");
            chartArea.AxisX.MajorGrid.LineWidth = 0;
            chartArea.AxisY.MajorGrid.LineWidth = 0;
            chartArea.BackColor = Color.White;

            chartArea.AxisY.Interval = 75;

            chart1.ChartAreas.Add(chartArea);
            chart1.BackColor = Color.White;
            chart1.Series.Clear();
            chart1.BorderColor = Color.White;
            chartArea.BorderColor = Color.White;

            chart1.ChartAreas[0].Area3DStyle.Enable3D = true;
            chart1.ChartAreas[0].Area3DStyle.LightStyle = LightStyle.Realistic;

            var series = new Series()
            {
                Name = "Averages",
                Color = Color.LightGray,
                ChartType = chartType
            };
            chart1.Series.Add(series);
            DataPoint dp = new DataPoint()
            {
                Color = Color.Salmon,
                BorderColor = Color.Salmon,
                MarkerBorderWidth = 10,
                AxisLabel = chartType != SeriesChartType.Pie ? String.Empty : "Sedentary",
                LabelForeColor = Color.Black,
                YValues = new double[] {SubjectRecords.SubjectValidAverages.AvgSedentary},
                BorderWidth = 1
            };
            dp.Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 7);
            chart1.Series[0].Points.Add(dp);
            dp = new DataPoint()
            {
                YValues = new double[] {SubjectRecords.SubjectValidAverages.AvgLight},
                Color = Color.DodgerBlue,
                BorderColor = Color.DodgerBlue,
                AxisLabel = chartType != SeriesChartType.Pie ? String.Empty : "Light",
                LabelForeColor = Color.Black,
                BorderWidth = 1
            };
            dp.Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 7);
            chart1.Series[0].Points.Add(dp);
            dp = new DataPoint()
            {
                YValues = new double[] {SubjectRecords.SubjectValidAverages.AvgLifestyle},
                Color = Color.GreenYellow,
                BorderColor = Color.GreenYellow,
                AxisLabel = chartType != SeriesChartType.Pie ? String.Empty : "Life Style",
                LabelForeColor = Color.Black,
                BorderWidth = 1
            };
            dp.Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 7);
            chart1.Series[0].Points.Add(dp);
            if (Math.Abs(SubjectRecords.SubjectValidAverages.AvgModerate10) > 0 && chartType == SeriesChartType.Pie)
            {
                dp = new DataPoint()
                {
                    YValues = new double[] {SubjectRecords.SubjectValidAverages.AvgModerate10},
                    Color = Color.Green,
                    LabelForeColor = Color.Black,
                    AxisLabel = chartType != SeriesChartType.Pie ? String.Empty : "Moderate-10",
                    BorderColor = Color.Black,
                    BorderDashStyle = ChartDashStyle.Dot,
                    BorderWidth = 1
                };
                dp.Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 7);
                chart1.Series[0].Points.Add(dp);
            }

            dp = new DataPoint()
            {
                Color = Color.Green,
                YValues =
                    new double[]
                    {
                        chartType == SeriesChartType.Pie
                            ? SubjectRecords.SubjectValidAverages.AvgModerate -
                              SubjectRecords.SubjectValidAverages.AvgModerate10
                            : SubjectRecords.SubjectValidAverages.AvgModerate
                    },
                AxisLabel = chartType != SeriesChartType.Pie ? String.Empty : "Moderate",
                LabelForeColor = Color.Black,
                BorderColor = Color.Green,
                BorderWidth = 1
            };
            dp.Font = new System.Drawing.Font(FontFamily.GenericSansSerif, 7);

            chart1.Series[0].Points.Add(dp);
            chart1.Series[0].IsValueShownAsLabel = true;

            if (chartType == SeriesChartType.Pie)
            {
                chart1.Series[0]["PieLabelStyle"] = "OutSide";
                chart1.Legends.Add(new Legend("legend1"));
                chart1.Legends[0].Alignment = StringAlignment.Far;
                chart1.Legends[0].Docking = Docking.Bottom;
                chart1.Legends[0].Enabled = true;
            }
            MemoryStream memoryStream = new MemoryStream();
            chart1.SaveImage(memoryStream, ChartImageFormat.Png);
            return memoryStream;
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
                Average_Steps_Valid_Days = Math.Round(avgSteps),
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
            Rectangle rec = new Rectangle(PageSize.LETTER);
            Document doc = new Document(rec, 36, 36, 36, 36);
            doc.SetMargins(20f, 36f, 18f, 18f);
            var fileName = GetFileName();
            //File.Copy(@"C:\Users\K.Vineel\Destop\TemplateDoc.pdf", fileName, true);
            FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            //fs.Close();
            PdfWriter writer = PdfWriter.GetInstance(doc, fs);
            doc.Open();
            Paragraph para = new Paragraph("Actigraph Summary Report");
            para.Font.SetStyle("bold");
            para.Alignment = Element.ALIGN_CENTER;
            doc.Add(para);
            para = new Paragraph("Report Date: " + DirectoryStructure.CurrentDateTime);
            para.Font.SetStyle("bold");
            para.Alignment = Element.ALIGN_RIGHT;
            doc.Add(para);
            para = new Paragraph("Dear Mr." + SubjectRecords.SubjectData.FirstOrDefault().FamilyName + ",");
            para.Font.SetStyle("bold");
            para.Alignment = Element.ALIGN_LEFT;
            doc.Add(para);
            //Table-1
            doc.Add(new Paragraph("\n"));
            para = new Paragraph("Subject Details");
            para.Font.SetStyle("bold");
            para.Font.SetStyle("underline");
            para.Alignment = Element.ALIGN_LEFT;
            doc.Add(para);
            doc.Add(new Paragraph("\n"));
            doc.Add(CreateTable(tableData1, "Subject Details"));
            para = new Paragraph($"Date Range: {tableData2.FirstOrDefault(t => t.Name == "Date Range").Value}");
            para.Alignment = Element.ALIGN_LEFT;
            doc.Add(para);
            doc.Add(new Paragraph("\n"));
            //Start Table-2
            //doc.Add(CreateTable(tableData2, "Subject Summary Details"));
            //Insert Graph
            para = new Paragraph("Graph Showing Subject Averages");
            para.Font.SetStyle("bold");
            para.Font.SetStyle("underline");
            para.Alignment = Element.ALIGN_LEFT;
            para.PaddingTop = 50f;
            doc.Add(para);
            PdfPTable table = new PdfPTable(2);
            table.PaddingTop = 50f;

            table.TotalWidth = 580f;
            table.LockedWidth = true;
            table.HorizontalAlignment = 0;
            float[] widths = {260, 320};
            table.SetWidths(widths);
            PdfPCell cell = new PdfPCell();
            table.DefaultCell.Border = Rectangle.NO_BORDER;
            Image barChart = Image.GetInstance(CreateGraph(SeriesChartType.StackedBar).ToArray());
            //barChart.ScaleToFit(250f,220f);
            barChart.Alignment = iTextSharp.text.Image.ALIGN_LEFT;
            cell = new PdfPCell(barChart);
            cell.HorizontalAlignment = 0;
            cell.Border = 0;
            table.AddCell(cell);
            cell = new PdfPCell();
            Image PieChart = Image.GetInstance(CreateGraph(SeriesChartType.Pie).ToArray());
            //PieChart.ScaleToFit(330f, 220f);
            PieChart.Alignment = iTextSharp.text.Image.ALIGN_LEFT;
            cell = new PdfPCell(PieChart);
            cell.HorizontalAlignment = 0;
            cell.Border = 0;
            table.AddCell(cell);
            doc.Add(table);
            doc.Add(new Paragraph("\n"));
            para = new Paragraph("Subject Summary Details (Weekly)");
            para.Font.SetStyle("bold");
            para.Font.SetStyle("underline");
            para.Alignment = Element.ALIGN_LEFT;
            para.PaddingTop = 50f;
            doc.Add(para);
            doc.Add(new Paragraph("\n"));
            doc.Add(CreateSummaryTable(tableData3));
            doc.Close();
        }

        private string GetFileName()
        {
            var subjectId = tableData1.Where(i => i.Name == "ID").ToArray()[0].Value;
            var folderName = DirectoryStructure.CreateSubjectFolder(subjectId);
            string fileName = Path.Combine(folderName,
                subjectId + "-" + DirectoryStructure.CurrentDateTime + fileExtension);
            if (File.Exists(fileName))
            {
                fileName = Path.Combine(folderName,
                    Guid.NewGuid() +
                    Path.GetExtension(fileName.ToString()));
            }
            return fileName;
        }

        private PdfPTable CreateTable(List<TableData> dataTable, string tableHeader)
        {
            var para = new Paragraph(tableHeader);
            para.Font.SetStyle("bold");
            para.Font.SetStyle("underline");
            para.Alignment = Element.ALIGN_LEFT;
            para.PaddingTop = 100f;
            PdfPTable table = new PdfPTable(dataTable.Count);
            table.PaddingTop = 100f;

            table.TotalWidth = 530f;
            table.LockedWidth = true;
            table.HorizontalAlignment = 0;
            //float[] widths = new float[] { 150f, 150f };
            //table.SetWidths(widths);
            PdfPCell cell = new PdfPCell();
            table.DefaultCell.Border = Rectangle.NO_BORDER;
            foreach (var item in dataTable)
            {
                cell = new PdfPCell(new Paragraph(item.Name));
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                //cell.Width = 150f;
                table.AddCell(cell);
                cell = new PdfPCell(new Paragraph(item.Value));
                cell.HorizontalAlignment = 0;
                cell.Border = 0;
                //cell.Width = 150f;
                table.AddCell(cell);
            }
            return table;
        }

        private PdfPTable CreateSittingSummaryTable(string tableHeader)
        {
            var para = new Paragraph(tableHeader);
            para.Font.SetStyle("bold");
            para.Font.SetStyle("underline");
            para.Alignment = Element.ALIGN_LEFT;
            para.PaddingTop = 100f;
            PdfPTable table = new PdfPTable(2);
            table.PaddingTop = 100f;

            table.TotalWidth = 530f;
            table.LockedWidth = true;
            table.HorizontalAlignment = 0;
            float[] widths = new float[] {150f, 150f};
            table.SetWidths(widths);
            PdfPCell cell = new PdfPCell(para);
            cell.Colspan = 2;
            cell.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
            cell.Border = 0;
            //cell.Width = 150f;
            cell.PaddingBottom = 20f;
            table.AddCell(cell);
            table.DefaultCell.Border = Rectangle.NO_BORDER;
            cell = new PdfPCell(new Paragraph("Average_Sitting_All_Days".Replace('_', ' ')));
            cell.HorizontalAlignment = 0;
            cell.Border = 0;
            table.AddCell(cell);
            cell =
                new PdfPCell(
                    new Paragraph(Math.Round(SubjectRecords.SubjectData.Select(p => p.Sedentary).Average()).ToString()));
            cell.HorizontalAlignment = 0;
            cell.Border = 0;
            table.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Average_Sitting_Valid_Days".Replace('_', ' ')));
            cell.HorizontalAlignment = 0;
            cell.Border = 0;
            table.AddCell(cell);
            cell = new PdfPCell(new Paragraph(Math.Round(SubjectRecords.SubjectValidAverages.AvgSedentary).ToString()));
            cell.HorizontalAlignment = 0;
            cell.Border = 0;
            table.AddCell(cell);
            return table;
        }

        private PdfPTable CreateSummaryTable(SummaryTableData[] summaryTable)
        {
            PdfPTable table = new PdfPTable(9);
            table.HorizontalAlignment = 0;
            table.PaddingTop = 100f;
            float[] widths = {230f, 240f, 200f, 205f, 160f, 170f, 170f, 170f, 220f};
            table.TotalWidth = widths.Sum(f => f)/3f;
            table.LockedWidth = true;
            table.SetWidths(widths);
            PdfPCell cell = new PdfPCell();
            cell.Colspan = 9;
            cell.HorizontalAlignment = 0; //0=Left, 1=Centre, 2=Right
            cell.Border = 0;
            //cell.Width = 150f;
            cell.PaddingBottom = 20f;
            table.AddCell(cell);
            table.DefaultCell.Border = Rectangle.NO_BORDER;
            //Cell-1

            table.AddCell(GetCell("Date", BaseColor.BLUE));
            table.AddCell(GetCell("Day of Week", BaseColor.BLUE));
            table.AddCell(GetCell("Wear Time", BaseColor.BLUE));
            table.AddCell(GetCell("Movements Per Minute", BaseColor.BLUE));
            table.AddCell(GetCell("Steps", BaseColor.BLUE));
            table.AddCell(GetCell("Light", BaseColor.BLUE));
            table.AddCell(GetCell("Life Style", BaseColor.BLUE));
            table.AddCell(GetCell("Moderate", BaseColor.BLUE));
            table.AddCell(GetCell("Moderate-10", BaseColor.BLUE));
            //Create data rows
            foreach (var data in summaryTable)
            {
                var color = data.Wear_Time < ThresholdCutOffValid ? BaseColor.LIGHT_GRAY : null;

                table.AddCell(GetCell(data.Date, color, true));
                table.AddCell(GetCell(data.Day_of_Week, color, true));
                table.AddCell(GetCell(data.Wear_Time.ToString(CultureInfo.InvariantCulture), color, true));
                table.AddCell(GetCell(data.Movements_Per_Minute.ToString(CultureInfo.InvariantCulture), color, true));
                table.AddCell(GetCell(data.Steps.ToString(CultureInfo.InvariantCulture), color, true));
                table.AddCell(GetCell(data.Light.ToString(CultureInfo.InvariantCulture), color, true));
                table.AddCell(GetCell(data.LifeStyle.ToString(CultureInfo.InvariantCulture), color, true));
                table.AddCell(GetCell(data.Moderate.ToString(CultureInfo.InvariantCulture), color, true));
                table.AddCell(GetCell(data.Moderate10.ToString(CultureInfo.InvariantCulture), color, true));
            }
            //Create Summary Last Row
            table.AddCell(GetCell("Avg's for valid days(" + SubjectRecords.SubjectValidAverages.ValidDays + ")"));
            table.AddCell(GetCell("N/A"));
            table.AddCell(GetCell(tableData2.Where(t => t.Name == "Average Wear Time Valid Days").ToArray()[0].Value));
            table.AddCell(GetCell(tableData2.Where(t => t.Name == "Average Movements Valid Days").ToArray()[0].Value));
            table.AddCell(GetCell(tableData2.Where(t => t.Name == "Average Steps Valid Days").ToArray()[0].Value));
            table.AddCell(GetCell(SubjectRecords.SubjectValidAverages.AvgLight.ToString(CultureInfo.InvariantCulture)));
            table.AddCell(
                GetCell(SubjectRecords.SubjectValidAverages.AvgLifestyle.ToString(CultureInfo.InvariantCulture)));
            table.AddCell(GetCell(SubjectRecords.SubjectValidAverages.AvgModerate.ToString(CultureInfo.InvariantCulture)));
            table.AddCell(
                GetCell(SubjectRecords.SubjectValidAverages.AvgModerate10.ToString(CultureInfo.InvariantCulture)));
            return table;
        }

        private PdfPCell GetCell(string cellText, BaseColor bgColor = null, bool isCellData = false)
        {
            BaseFont bfTimes = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
            Font times = new Font(bfTimes, 12, Font.NORMAL, BaseColor.WHITE);
            if (isCellData)
                times = new Font(bfTimes, 12, Font.NORMAL, BaseColor.BLACK);
            var cell = new PdfPCell();
            cell = bgColor != null
                ? new PdfPCell(new Paragraph(cellText, times))
                : new PdfPCell(new Paragraph(cellText));
            cell.HorizontalAlignment = bgColor != null && !isCellData ? 1 : 0;
            cell.BackgroundColor = bgColor ?? BaseColor.WHITE;
            return cell;
        }
    }
}
