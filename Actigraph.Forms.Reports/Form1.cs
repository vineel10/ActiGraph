using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
                lbldateTime.Text = DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt");
                button1.TabStop = false;
                button1.FlatStyle = FlatStyle.Flat;
                button1.FlatAppearance.BorderSize = 0;
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
                var stagedFiles = DirectoryStructure.GetStagedFiles();
                lblStagedFiles.Text = string.Format("Currently you have {0} files to process.",
                    stagedFiles.Count());
                listView1.View = View.Details;
                //listView1.GridLines = true;
                listView1.FullRowSelect = true;
                listView1.Sort();
                //listView1.CheckBoxes = true;
                //Add column header
                listView1.Columns.Add("FileName", 300);
                // listView1.Items.AddRange(stagedFiles);
                foreach (var itm in stagedFiles.Select(item => new ListViewItem(item)))
                {
                    listView1.Items.Add(Path.GetFileName(itm.Text));
                }
            }
            catch (Exception exp)
            {
                log.Error(exp);
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
            try
            {
                if (listView1.SelectedItems.Count < 1)
                {
                    MessageBox.Show("Please select atleast 1 file to process.");
                    return;
                }
                //MessageBox.Show(listView1.SelectedItems[0].SubItems[0].Text);
                XlFileProcessor processor = new XlFileProcessor();
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
                        CreateDocFiles files = new CreateDocFiles();
                        files.fileExtension = ".PDF";
                        files.CreateReports(item);
                        files = null;
                        processedBar.Value = i;
                        lblProcessing.Text = "Processing " + i + " of " + processedBar.Maximum;
                        Application.DoEvents();
                        i++;
                        
                    }
                    processedBar.Visible = false;
                    lblProcessing.Visible = false;
                    MessageBox.Show("Reports generated Successfuuly!!!. Please press Ok to open reports folder.");
                    Process.Start(DirectoryStructure.ActigraphReportsFolderName);
                }
                else
                {
                    MessageBox.Show("Invalid data in selected file");
                }
            }
            catch (Exception exp)
            {
                button1.Enabled = true;
                MessageBox.Show(exp.Message);
                log.Error(exp);
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            GenerateReports();
            DirectoryStructure.MoveProcessedFile(listView1.SelectedItems[0].SubItems[0].Text);
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
    }
}
