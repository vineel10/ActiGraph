namespace Actigraph.Forms.Reports
{
    partial class LoadData
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.label1 = new System.Windows.Forms.Label();
            this.lbldateTime = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.lblStagedFiles = new System.Windows.Forms.Label();
            this.processedBar = new System.Windows.Forms.ProgressBar();
            this.lblProcessing = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.label1.Location = new System.Drawing.Point(86, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(377, 40);
            this.label1.TabIndex = 0;
            this.label1.Text = "Welcome to Actigraph Reports Extractor";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // lbldateTime
            // 
            this.lbldateTime.ForeColor = System.Drawing.SystemColors.ControlLightLight;
            this.lbldateTime.Location = new System.Drawing.Point(191, 408);
            this.lbldateTime.Name = "lbldateTime";
            this.lbldateTime.Size = new System.Drawing.Size(133, 23);
            this.lbldateTime.TabIndex = 1;
            this.lbldateTime.Text = "DateTime";
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.BackColor = System.Drawing.Color.White;
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.lblProcessing);
            this.panel1.Controls.Add(this.processedBar);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.listView1);
            this.panel1.Controls.Add(this.lblStagedFiles);
            this.panel1.Location = new System.Drawing.Point(29, 60);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(489, 332);
            this.panel1.TabIndex = 2;
            // 
            // button1
            // 
            this.button1.AutoSize = true;
            this.button1.BackColor = System.Drawing.Color.MidnightBlue;
            this.button1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.Color.White;
            this.button1.Location = new System.Drawing.Point(19, 190);
            this.button1.Margin = new System.Windows.Forms.Padding(0);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(209, 34);
            this.button1.TabIndex = 2;
            this.button1.Text = "Generate Reports";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // listView1
            // 
            this.listView1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listView1.FullRowSelect = true;
            this.listView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.listView1.Location = new System.Drawing.Point(85, 59);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(321, 72);
            this.listView1.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            // 
            // lblStagedFiles
            // 
            this.lblStagedFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStagedFiles.Location = new System.Drawing.Point(7, 12);
            this.lblStagedFiles.Name = "lblStagedFiles";
            this.lblStagedFiles.Size = new System.Drawing.Size(428, 24);
            this.lblStagedFiles.TabIndex = 0;
            // 
            // processedBar
            // 
            this.processedBar.AccessibleDescription = "";
            this.processedBar.BackColor = System.Drawing.Color.Lime;
            this.processedBar.ForeColor = System.Drawing.Color.White;
            this.processedBar.Location = new System.Drawing.Point(122, 258);
            this.processedBar.Name = "processedBar";
            this.processedBar.Size = new System.Drawing.Size(226, 23);
            this.processedBar.TabIndex = 3;
            this.processedBar.Visible = false;
            // 
            // lblProcessing
            // 
            this.lblProcessing.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblProcessing.ForeColor = System.Drawing.Color.MidnightBlue;
            this.lblProcessing.Location = new System.Drawing.Point(148, 284);
            this.lblProcessing.Name = "lblProcessing";
            this.lblProcessing.Size = new System.Drawing.Size(159, 23);
            this.lblProcessing.TabIndex = 4;
            this.lblProcessing.Visible = false;
            // 
            // btnCancel
            // 
            this.btnCancel.AutoSize = true;
            this.btnCancel.BackColor = System.Drawing.Color.MidnightBlue;
            this.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Location = new System.Drawing.Point(260, 190);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(0);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(209, 34);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.button2_Click);
            // 
            // LoadData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.MidnightBlue;
            this.ClientSize = new System.Drawing.Size(547, 430);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lbldateTime);
            this.Controls.Add(this.label1);
            this.Name = "LoadData";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.LoadData_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lbldateTime;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblStagedFiles;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ProgressBar processedBar;
        private System.Windows.Forms.Label lblProcessing;
        private System.Windows.Forms.Button btnCancel;
    }
}

