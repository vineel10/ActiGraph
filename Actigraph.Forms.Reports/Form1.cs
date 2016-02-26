using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Actigraph.Forms.Reports.Properties;
using Actigraph.Parser;
using Actigraph.Parser.Generate_DocFiles;

namespace Actigraph.Forms.Reports
{
    public partial class LoadData : Form
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger
            (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public LoadData()
        {
            try
            {
                var color = Color.FromArgb(4, 19, 34);
                InitializeComponent();
                Text = "Load Data";
                BackColor = color;
                panel1.BackColor = color;
                lblStagedFiles.ForeColor = Color.White;
                lblProcessing.ForeColor = Color.White;
                listView1.BackColor = color;
                listView1.ForeColor = Color.White;
                listView1.Scrollable = false;
                listView1.BorderStyle = BorderStyle.FixedSingle;
                listView1.HeaderStyle = ColumnHeaderStyle.None;
                listView1.MultiSelect = false;

                label2.ForeColor = Color.White;
                lbldateTime.Text = DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt");
                txtThresholdCutoff.Text = "600";
                ChangeButtons();
            }
            catch (Exception exp)
            {
                log.Error(exp);
                MessageBox.Show(exp.Message);
            }
        }

        private void ChangeButtons()
        {
            var color = Color.FromArgb(4, 19, 34);
            foreach (var control in this.panel1.Controls)
            {
                if (typeof (Button).FullName == control.GetType().FullName)
                {
                    ((Button) control).BackColor = color;
                    ((Button) control).TabStop = false;
                    ((Button) control).FlatStyle = FlatStyle.Flat;
                    ((Button) control).FlatAppearance.BorderSize = 1;
                    ((Button) control).FlatAppearance.BorderColor = Color.White;
                }
            }
        }

        private void LoadData_Load(object sender, EventArgs e)
        {
            try
            {
                DirectoryStructure.CreateDirectoryStructure();
                log.Info("Folder structure created.");
                CreateList();
            }
            catch (Exception exp)
            {
                log.Error(exp);
                MessageBox.Show(exp.Message);
            }
        }

        private void CreateList()
        {
            try
            {
                var stagedFiles = DirectoryStructure.GetStagedFiles();
                lblStagedFiles.Text = $"Currently you have {stagedFiles.Count()} files to process.";
                listView1.Columns.Clear();
                listView1.Items.Clear();
                listView1.View = View.Details;
                //listView1.GridLines = true;
                listView1.FullRowSelect = true;
                listView1.Sort();
                //listView1.CheckBoxes = true;
                //Add column header

                // listView1.Items.AddRange(stagedFiles);
                foreach (var itm in stagedFiles.Select(item => new ListViewItem(item)))
                {
                    listView1.Items.Add(Path.GetFileName(itm.Text));
                }
                if (stagedFiles.Length > 0)
                {
                    listView1.Columns.Add("File Name", 300);
                    listView1.Columns[0].AutoResize(ColumnHeaderAutoResizeStyle.ColumnContent);
                }
            }
            catch (Exception exp)
            {
                log.Error(exp);
                MessageBox.Show(exp.Message);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                lbldateTime.Text = DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt");
                Application.DoEvents();
            }
            catch (Exception exp)
            {
                log.Error(exp);
                MessageBox.Show(exp.Message);
            }
        }


        private bool GenerateReports()
        {
            long ThresholdCutoff = 0L;
            try
            {
                if (listView1.SelectedItems.Count < 1)
                {
                    MessageBox.Show(@"Please select a file to process.");
                    button1.Enabled = true;
                    return false;
                }
                if (!long.TryParse(txtThresholdCutoff.Text, out ThresholdCutoff))
                {
                    MessageBox.Show(@"Enter valid CutOff");
                    return false;
                }
                txtThresholdCutoff.Enabled = false;
                //MessageBox.Show(listView1.SelectedItems[0].SubItems[0].Text);
                XlFileProcessor processor = new XlFileProcessor(ThresholdCutoff);
                var parsedSubjectsData = processor.LoadFile(listView1.SelectedItems[0].SubItems[0].Text);
                if (parsedSubjectsData != null)
                {
                    int i = 1;
                    processedBar.Minimum = 0;
                    var subjectRecordses = parsedSubjectsData as SubjectRecords[] ?? parsedSubjectsData.ToArray();
                    processedBar.Maximum = subjectRecordses.Count();
                    processedBar.Style = ProgressBarStyle.Continuous;
                    processedBar.TabIndex = 0;
                    processedBar.Visible = true;
                    lblProcessing.Visible = true;
                    lblProcessing.Text = $"Processing {i} of {processedBar.Maximum}";
                    foreach (var item in subjectRecordses)
                    {
                        //CreateDocFiles files = new CreateDocFiles();
                        CreatePdfReports files = new CreatePdfReports(ThresholdCutoff)
                        {fileExtension = @".pdf"};
                        files.CreateReports(item);
                        files = null;
                        processedBar.Value = i;
                        lblProcessing.Text = $"Processing {i} of {processedBar.Maximum}";
                        Application.DoEvents();
                        i++;
                    }
                }
                else
                {
                    MessageBox.Show(@"Invalid data in selected file");
                    button1.Enabled = true;
                    txtThresholdCutoff.Enabled = true;
                    return false;
                }
            }
            catch (Exception exp)
            {
                log.Error(exp);
                button1.Enabled = true;
                txtThresholdCutoff.Enabled = true;
                MessageBox.Show(exp.Message);
                return false;
            }
            return true;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                button1.Enabled = false;
                button1.ForeColor = Color.White;
                if (!GenerateReports()) return;
                DirectoryStructure.MoveProcessedFile(listView1.SelectedItems[0].SubItems[0].Text);
                processedBar.Visible = false;
                lblProcessing.Visible = false;
                CreateList();
                button1.Enabled = true;
                txtThresholdCutoff.Enabled = true;
                MessageBox.Show(Resources.LoadData_button1_Click_Reports_generated_Successfuuly);
                Process.Start(Path.Combine(DirectoryStructure.ActigraphReportsFolderName,
                    DateTime.Now.ToShortDateString()));
            }
            catch (Exception exp)
            {
                log.Error(exp);
                MessageBox.Show(exp.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(Resources.LoadData_button2_Click_Exit_application_,
                Resources.LoadData_button2_Click_Close_Application,
                MessageBoxButtons.OKCancel) != DialogResult.Cancel)
            {
                Application.ExitThread();
                Application.Exit();
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            CreateList();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();

                openFileDialog1.InitialDirectory = @"C:\";
                openFileDialog1.Title = @"Browse Excel Files";

                openFileDialog1.CheckFileExists = true;
                openFileDialog1.CheckPathExists = true;

                openFileDialog1.DefaultExt = "xlsx";
                openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;";
                openFileDialog1.RestoreDirectory = true;

                openFileDialog1.Multiselect = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    textBox2.Text = openFileDialog1.FileNames.Aggregate((current, next) => current + ", " + next);
                    foreach (var file in openFileDialog1.FileNames)
                    {
                        var extension = Path.GetExtension(file);
                        if (extension == null || !extension.ToUpperInvariant().Contains("XLS")) continue;
                        var fileName = Path.GetFileName(file);
                        File.Copy(file, Path.Combine(DirectoryStructure.ActigraphDataFilesFolderName, fileName), true);
                    }
                    CreateList();
                }
            }
            catch (Exception exp)
            {
                log.Error(exp);
                MessageBox.Show(exp.Message);
            }
        }
    }
}
