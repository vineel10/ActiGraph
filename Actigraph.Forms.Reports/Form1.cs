using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
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
                InitializeComponent();
                this.Text = "Load Data";
                this.BackColor = Color.FromArgb(0, 113, 113);
                button1.BackColor= Color.FromArgb(0, 113, 113);
                btnCancel.BackColor= Color.FromArgb(0, 113, 113);
                button2.BackColor = Color.FromArgb(0, 113, 113);
                lbldateTime.Text = DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt");
                txtThresholdCutoff.Text = "600";
                button1.TabStop = false;
                button1.FlatStyle = FlatStyle.Flat;
                button1.FlatAppearance.BorderSize = 0;
                button2.TabStop = false;
                button2.FlatStyle = FlatStyle.Flat;
                button2.FlatAppearance.BorderSize = 0;
                btnCancel.TabStop = false;
                btnCancel.FlatStyle = FlatStyle.Flat;
                btnCancel.FlatAppearance.BorderSize = 0;
            }
            catch (Exception exp)
            {
                log.Error(exp);
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
            }
        }

        private void CreateList()
        {
            var stagedFiles = DirectoryStructure.GetStagedFiles();
            lblStagedFiles.Text = string.Format("Currently you have {0} files to process.",
                stagedFiles.Count());
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
            }
        }


        private void GenerateReports()
        {
            long ThresholdCutoff = 0L;
            try
            {
                if (listView1.SelectedItems.Count < 1)
                {
                    MessageBox.Show("Please select atleast 1 file to process.");
                    button1.Enabled = true;
                    return;
                }
                if (!long.TryParse(txtThresholdCutoff.Text, out ThresholdCutoff))
                {
                    MessageBox.Show("Enter valid CutOff");
                    return;
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
                    lblProcessing.Text = "Processing 1" + " of " + processedBar.Maximum;
                    foreach (var item in subjectRecordses)
                    {
                        //CreateDocFiles files = new CreateDocFiles();
                        CreatePdfReports files = new CreatePdfReports(ThresholdCutoff);
                        files.fileExtension = ".pdf";
                        files.CreateReports(item);
                        files = null;
                        processedBar.Value = i;
                        lblProcessing.Text = "Processing " + i + " of " + processedBar.Maximum;
                        Application.DoEvents();
                        i++;
                    }
                    
                }
                else
                {
                    MessageBox.Show("Invalid data in selected file");
                    button1.Enabled = true;
                    txtThresholdCutoff.Enabled = true;
                }
            }
            catch (Exception exp)
            {
                button1.Enabled = true;
                txtThresholdCutoff.Enabled = true;
                MessageBox.Show(exp.Message);
                log.Error(exp);
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {

            button1.Enabled = false;
            GenerateReports();
            DirectoryStructure.MoveProcessedFile(listView1.SelectedItems[0].SubItems[0].Text);
            processedBar.Visible = false;
            lblProcessing.Visible = false;
            CreateList();
            button1.Enabled = true;
            txtThresholdCutoff.Enabled = true;
            MessageBox.Show("Reports generated Successfuuly!!!. Please press Ok to open reports folder.");
            Process.Start(DirectoryStructure.ActigraphReportsFolderName);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Exit application?", "Close Application",
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
    }
}
