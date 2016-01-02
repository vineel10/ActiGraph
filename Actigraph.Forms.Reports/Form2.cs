using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace Actigraph.Forms.Reports
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            Chart chart1 = new Chart()
            {

                Width = 700,
                Height = 300
            };

            ChartArea chartArea = new ChartArea("chartArea1");
            chartArea.AxisX.MajorGrid.LineWidth = 0;
            chartArea.AxisY.MajorGrid.LineWidth = 0;
            chartArea.BackColor = Color.Goldenrod;
            
            chartArea.AxisY.Interval = 75;
            chartArea.InnerPlotPosition = new ElementPosition()
            {
                Height = 80,
                Width = 90,
                X = 15,
                Y = 5
            };
            chart1.ChartAreas.Add(chartArea);
            chart1.ChartAreas[0].Area3DStyle.Enable3D = true;
            chart1.BackColor = Color.Transparent;
            chart1.Series.Clear();
            var series = new Series()
            {
                Name = "Averages",
                Color = Color.LightGray,
                ChartType = SeriesChartType.StackedBar
            };
            chart1.Series.Add(series);
            DataPoint dp = new DataPoint()
            {
                Color = Color.DarkSeaGreen,
                YValues = new double[] {612.79},
                AxisLabel = "Sedentary",
                Label = "612.79"

            };
            chart1.Series[0].Points.Add(dp);
            dp = new DataPoint()
            {
                Color = Color.LightGreen,
                YValues = new double[] { 111 },
                AxisLabel = "Light",
                Label = "111"
            };
            chart1.Series[0].Points.Add(dp);
            dp = new DataPoint()
            {
                Color = Color.LawnGreen,
                YValues = new double[] { 65 },
                AxisLabel = "Life Style",
                Label = "65"
            };
            chart1.Series[0].Points.Add(dp);
            dp = new DataPoint()
            {
                Color = Color.DarkGreen,
                YValues = new double[] { 27 },
                AxisLabel = "Moderate",
                Label = "27"
            };
            chart1.Series[0].Points.Add(dp);
            this.Controls.Add(chart1);
        }
    }
}
