namespace LinkSourceChanger
{
    partial class MainPage
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainPage));
            this.folderBrowser = new System.Windows.Forms.FolderBrowserDialog();
            this.buttonTargetFolder = new System.Windows.Forms.Button();
            this.labelTargetFolder = new System.Windows.Forms.Label();
            this.buttonGetLinks = new System.Windows.Forms.Button();
            this.textBoxLog = new System.Windows.Forms.TextBox();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.labelProgLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.labelProcessStatus = new System.Windows.Forms.ToolStripStatusLabel();
            this.progBar = new System.Windows.Forms.ToolStripProgressBar();
            this.labelLinkSources = new System.Windows.Forms.Label();
            this.listLinks = new System.Windows.Forms.ListView();
            this.currentLinkSourcePath = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.newLinkSourcePath = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.replaceSource = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.replaceTarget = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.textReplaceSource = new System.Windows.Forms.TextBox();
            this.textReplaceDestination = new System.Windows.Forms.TextBox();
            this.labelReplaceSource = new System.Windows.Forms.Label();
            this.labelReplaceDestination = new System.Windows.Forms.Label();
            this.buttonReplaceAllLinkSources = new System.Windows.Forms.Button();
            this.buttonSetNewDesitnation = new System.Windows.Forms.Button();
            this.panelMain = new System.Windows.Forms.SplitContainer();
            this.statusStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelMain)).BeginInit();
            this.panelMain.Panel1.SuspendLayout();
            this.panelMain.Panel2.SuspendLayout();
            this.panelMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // folderBrowser
            // 
            this.folderBrowser.RootFolder = System.Environment.SpecialFolder.MyComputer;
            // 
            // buttonTargetFolder
            // 
            this.buttonTargetFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.buttonTargetFolder.Location = new System.Drawing.Point(12, 40);
            this.buttonTargetFolder.Name = "buttonTargetFolder";
            this.buttonTargetFolder.Size = new System.Drawing.Size(95, 23);
            this.buttonTargetFolder.TabIndex = 0;
            this.buttonTargetFolder.Text = "Target Folder";
            this.buttonTargetFolder.UseVisualStyleBackColor = true;
            this.buttonTargetFolder.Click += new System.EventHandler(this.buttonTargetFolder_Click);
            // 
            // labelTargetFolder
            // 
            this.labelTargetFolder.AutoSize = true;
            this.labelTargetFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.labelTargetFolder.Location = new System.Drawing.Point(12, 9);
            this.labelTargetFolder.Name = "labelTargetFolder";
            this.labelTargetFolder.Size = new System.Drawing.Size(96, 13);
            this.labelTargetFolder.TabIndex = 1;
            this.labelTargetFolder.Text = "Select target folder";
            // 
            // buttonGetLinks
            // 
            this.buttonGetLinks.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.buttonGetLinks.Location = new System.Drawing.Point(113, 40);
            this.buttonGetLinks.Name = "buttonGetLinks";
            this.buttonGetLinks.Size = new System.Drawing.Size(94, 23);
            this.buttonGetLinks.TabIndex = 1;
            this.buttonGetLinks.Text = "Get Links";
            this.buttonGetLinks.UseVisualStyleBackColor = true;
            this.buttonGetLinks.Click += new System.EventHandler(this.buttonGetLinks_Click);
            // 
            // textBoxLog
            // 
            this.textBoxLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxLog.Location = new System.Drawing.Point(0, 0);
            this.textBoxLog.Multiline = true;
            this.textBoxLog.Name = "textBoxLog";
            this.textBoxLog.ReadOnly = true;
            this.textBoxLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxLog.Size = new System.Drawing.Size(1545, 279);
            this.textBoxLog.TabIndex = 7;
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.labelProgLabel,
            this.labelProcessStatus,
            this.progBar});
            this.statusStrip.Location = new System.Drawing.Point(0, 679);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(1569, 22);
            this.statusStrip.SizingGrip = false;
            this.statusStrip.TabIndex = 7;
            this.statusStrip.Text = "statusStrip1";
            // 
            // labelProgLabel
            // 
            this.labelProgLabel.Name = "labelProgLabel";
            this.labelProgLabel.Size = new System.Drawing.Size(60, 17);
            this.labelProgLabel.Text = "(progress)";
            // 
            // labelProcessStatus
            // 
            this.labelProcessStatus.Name = "labelProcessStatus";
            this.labelProcessStatus.Size = new System.Drawing.Size(1311, 17);
            this.labelProcessStatus.Spring = true;
            this.labelProcessStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // progBar
            // 
            this.progBar.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.progBar.Name = "progBar";
            this.progBar.Size = new System.Drawing.Size(150, 16);
            this.progBar.Step = 1;
            // 
            // labelLinkSources
            // 
            this.labelLinkSources.AutoSize = true;
            this.labelLinkSources.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.labelLinkSources.Location = new System.Drawing.Point(12, 94);
            this.labelLinkSources.Name = "labelLinkSources";
            this.labelLinkSources.Size = new System.Drawing.Size(70, 13);
            this.labelLinkSources.TabIndex = 9;
            this.labelLinkSources.Text = "Link sources:";
            // 
            // listLinks
            // 
            this.listLinks.Activation = System.Windows.Forms.ItemActivation.OneClick;
            this.listLinks.Alignment = System.Windows.Forms.ListViewAlignment.SnapToGrid;
            this.listLinks.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.listLinks.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.currentLinkSourcePath,
            this.newLinkSourcePath,
            this.replaceSource,
            this.replaceTarget});
            this.listLinks.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listLinks.GridLines = true;
            this.listLinks.Location = new System.Drawing.Point(0, 0);
            this.listLinks.MultiSelect = false;
            this.listLinks.Name = "listLinks";
            this.listLinks.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.listLinks.Size = new System.Drawing.Size(1545, 283);
            this.listLinks.TabIndex = 6;
            this.listLinks.UseCompatibleStateImageBehavior = false;
            this.listLinks.View = System.Windows.Forms.View.Details;
            this.listLinks.Click += new System.EventHandler(this.listLinks_Click);
            this.listLinks.Resize += new System.EventHandler(this.listLinks_Resize);
            // 
            // currentLinkSourcePath
            // 
            this.currentLinkSourcePath.Text = "Current link source path";
            this.currentLinkSourcePath.Width = 380;
            // 
            // newLinkSourcePath
            // 
            this.newLinkSourcePath.Text = "Target link source path";
            this.newLinkSourcePath.Width = 342;
            // 
            // replaceSource
            // 
            this.replaceSource.Text = "Replace source";
            this.replaceSource.Width = 485;
            // 
            // replaceTarget
            // 
            this.replaceTarget.Text = "Replace target";
            this.replaceTarget.Width = 836;
            // 
            // textReplaceSource
            // 
            this.textReplaceSource.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textReplaceSource.Location = new System.Drawing.Point(371, 13);
            this.textReplaceSource.Name = "textReplaceSource";
            this.textReplaceSource.Size = new System.Drawing.Size(1186, 20);
            this.textReplaceSource.TabIndex = 2;
            this.textReplaceSource.TextChanged += new System.EventHandler(this.textReplaceSource_TextChanged);
            // 
            // textReplaceDestination
            // 
            this.textReplaceDestination.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textReplaceDestination.Location = new System.Drawing.Point(371, 43);
            this.textReplaceDestination.Name = "textReplaceDestination";
            this.textReplaceDestination.Size = new System.Drawing.Size(1186, 20);
            this.textReplaceDestination.TabIndex = 3;
            this.textReplaceDestination.TextChanged += new System.EventHandler(this.textReplaceDestination_TextChanged);
            // 
            // labelReplaceSource
            // 
            this.labelReplaceSource.AutoSize = true;
            this.labelReplaceSource.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.labelReplaceSource.Location = new System.Drawing.Point(280, 16);
            this.labelReplaceSource.Name = "labelReplaceSource";
            this.labelReplaceSource.Size = new System.Drawing.Size(85, 13);
            this.labelReplaceSource.TabIndex = 13;
            this.labelReplaceSource.Text = "Replace source:";
            // 
            // labelReplaceDestination
            // 
            this.labelReplaceDestination.AutoSize = true;
            this.labelReplaceDestination.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.labelReplaceDestination.Location = new System.Drawing.Point(280, 46);
            this.labelReplaceDestination.Name = "labelReplaceDestination";
            this.labelReplaceDestination.Size = new System.Drawing.Size(80, 13);
            this.labelReplaceDestination.TabIndex = 14;
            this.labelReplaceDestination.Text = "Replace target:";
            // 
            // buttonReplaceAllLinkSources
            // 
            this.buttonReplaceAllLinkSources.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonReplaceAllLinkSources.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.buttonReplaceAllLinkSources.Location = new System.Drawing.Point(1390, 74);
            this.buttonReplaceAllLinkSources.Name = "buttonReplaceAllLinkSources";
            this.buttonReplaceAllLinkSources.Size = new System.Drawing.Size(167, 23);
            this.buttonReplaceAllLinkSources.TabIndex = 5;
            this.buttonReplaceAllLinkSources.Text = "Replace all link sources";
            this.buttonReplaceAllLinkSources.UseVisualStyleBackColor = true;
            this.buttonReplaceAllLinkSources.Click += new System.EventHandler(this.buttonReplaceAllLinkSources_Click);
            // 
            // buttonSetNewDesitnation
            // 
            this.buttonSetNewDesitnation.Location = new System.Drawing.Point(371, 74);
            this.buttonSetNewDesitnation.Name = "buttonSetNewDesitnation";
            this.buttonSetNewDesitnation.Size = new System.Drawing.Size(127, 23);
            this.buttonSetNewDesitnation.TabIndex = 4;
            this.buttonSetNewDesitnation.Text = "Update target list";
            this.buttonSetNewDesitnation.UseVisualStyleBackColor = true;
            this.buttonSetNewDesitnation.Click += new System.EventHandler(this.buttonSetNewDesitnation_Click);
            // 
            // panelMain
            // 
            this.panelMain.Location = new System.Drawing.Point(12, 110);
            this.panelMain.Name = "panelMain";
            this.panelMain.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // panelMain.Panel1
            // 
            this.panelMain.Panel1.Controls.Add(this.listLinks);
            // 
            // panelMain.Panel2
            // 
            this.panelMain.Panel2.Controls.Add(this.textBoxLog);
            this.panelMain.Size = new System.Drawing.Size(1545, 566);
            this.panelMain.SplitterDistance = 283;
            this.panelMain.TabIndex = 17;
            // 
            // MainPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1569, 701);
            this.Controls.Add(this.panelMain);
            this.Controls.Add(this.buttonSetNewDesitnation);
            this.Controls.Add(this.buttonReplaceAllLinkSources);
            this.Controls.Add(this.labelReplaceDestination);
            this.Controls.Add(this.labelReplaceSource);
            this.Controls.Add(this.textReplaceDestination);
            this.Controls.Add(this.textReplaceSource);
            this.Controls.Add(this.labelLinkSources);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.buttonGetLinks);
            this.Controls.Add(this.labelTargetFolder);
            this.Controls.Add(this.buttonTargetFolder);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(700, 350);
            this.Name = "MainPage";
            this.Text = "Change Excel Linksources";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainPage_FormClosing);
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.panelMain.Panel1.ResumeLayout(false);
            this.panelMain.Panel2.ResumeLayout(false);
            this.panelMain.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelMain)).EndInit();
            this.panelMain.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowser;
        private System.Windows.Forms.Button buttonTargetFolder;
        private System.Windows.Forms.Label labelTargetFolder;
        private System.Windows.Forms.Button buttonGetLinks;
        internal System.Windows.Forms.TextBox textBoxLog;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel labelProgLabel;
        private System.Windows.Forms.ToolStripStatusLabel labelProcessStatus;
        private System.Windows.Forms.ToolStripProgressBar progBar;
        private System.Windows.Forms.Label labelLinkSources;
        private System.Windows.Forms.ListView listLinks;
        private System.Windows.Forms.ColumnHeader currentLinkSourcePath;
        private System.Windows.Forms.ColumnHeader newLinkSourcePath;
        private System.Windows.Forms.TextBox textReplaceSource;
        private System.Windows.Forms.TextBox textReplaceDestination;
        private System.Windows.Forms.Label labelReplaceSource;
        private System.Windows.Forms.Label labelReplaceDestination;
        private System.Windows.Forms.Button buttonReplaceAllLinkSources;
        private System.Windows.Forms.Button buttonSetNewDesitnation;
        private System.Windows.Forms.ColumnHeader replaceSource;
        private System.Windows.Forms.ColumnHeader replaceTarget;
        private System.Windows.Forms.SplitContainer panelMain;
    }
}

