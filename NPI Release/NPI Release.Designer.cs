namespace NPI_Release
{
    partial class NPI_Release
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NPI_Release));
            this.VaultsLabel = new System.Windows.Forms.Label();
            this.VaultsComboBox = new System.Windows.Forms.ComboBox();
            this.OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnAccept = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tbECONumber = new System.Windows.Forms.TextBox();
            this.ECONumber = new System.Windows.Forms.Label();
            this.tbRevisionDescription = new System.Windows.Forms.TextBox();
            this.RevisionDescription = new System.Windows.Forms.Label();
            this.tbBatchRef = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.getRefsWorker = new System.ComponentModel.BackgroundWorker();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.CheckoutFilesWorker = new System.ComponentModel.BackgroundWorker();
            this.CheckinFilesWorker = new System.ComponentModel.BackgroundWorker();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.filesListView = new System.Windows.Forms.ListView();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.filesToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tableLayoutPanel1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // VaultsLabel
            // 
            this.VaultsLabel.AutoSize = true;
            this.VaultsLabel.Location = new System.Drawing.Point(6, 32);
            this.VaultsLabel.Name = "VaultsLabel";
            this.VaultsLabel.Size = new System.Drawing.Size(94, 13);
            this.VaultsLabel.TabIndex = 1;
            this.VaultsLabel.Text = " Select vault view:";
            // 
            // VaultsComboBox
            // 
            this.VaultsComboBox.FormattingEnabled = true;
            this.VaultsComboBox.Location = new System.Drawing.Point(12, 48);
            this.VaultsComboBox.Name = "VaultsComboBox";
            this.VaultsComboBox.Size = new System.Drawing.Size(126, 21);
            this.VaultsComboBox.TabIndex = 5;
            // 
            // OpenFileDialog
            // 
            this.OpenFileDialog.Title = "Open";
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(345, 3);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(96, 23);
            this.btnCancel.TabIndex = 6;
            this.btnCancel.Text = "Close";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // btnAccept
            // 
            this.btnAccept.Location = new System.Drawing.Point(247, 3);
            this.btnAccept.Name = "btnAccept";
            this.btnAccept.Size = new System.Drawing.Size(92, 23);
            this.btnAccept.TabIndex = 7;
            this.btnAccept.Text = "OK";
            this.btnAccept.UseVisualStyleBackColor = true;
            this.btnAccept.Click += new System.EventHandler(this.OKButton_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 98F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 102F));
            this.tableLayoutPanel1.Controls.Add(this.btnAccept, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnCancel, 3, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(9, 496);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(444, 30);
            this.tableLayoutPanel1.TabIndex = 9;
            // 
            // tbECONumber
            // 
            this.tbECONumber.Location = new System.Drawing.Point(12, 136);
            this.tbECONumber.Name = "tbECONumber";
            this.tbECONumber.Size = new System.Drawing.Size(112, 20);
            this.tbECONumber.TabIndex = 10;
            // 
            // ECONumber
            // 
            this.ECONumber.AutoSize = true;
            this.ECONumber.Location = new System.Drawing.Point(9, 120);
            this.ECONumber.Name = "ECONumber";
            this.ECONumber.Size = new System.Drawing.Size(69, 13);
            this.ECONumber.TabIndex = 11;
            this.ECONumber.Text = "ECO Number";
            // 
            // tbRevisionDescription
            // 
            this.tbRevisionDescription.Location = new System.Drawing.Point(130, 135);
            this.tbRevisionDescription.Name = "tbRevisionDescription";
            this.tbRevisionDescription.Size = new System.Drawing.Size(323, 20);
            this.tbRevisionDescription.TabIndex = 12;
            this.tbRevisionDescription.Text = "Release to Production";
            // 
            // RevisionDescription
            // 
            this.RevisionDescription.AutoSize = true;
            this.RevisionDescription.Location = new System.Drawing.Point(115, 120);
            this.RevisionDescription.Name = "RevisionDescription";
            this.RevisionDescription.Size = new System.Drawing.Size(104, 13);
            this.RevisionDescription.TabIndex = 13;
            this.RevisionDescription.Text = "Revision Description";
            // 
            // tbBatchRef
            // 
            this.tbBatchRef.Location = new System.Drawing.Point(12, 94);
            this.tbBatchRef.Name = "tbBatchRef";
            this.tbBatchRef.Size = new System.Drawing.Size(166, 20);
            this.tbBatchRef.TabIndex = 14;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(129, 13);
            this.label1.TabIndex = 15;
            this.label1.Text = "Finished Good to Release";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "AssemblyIcon.png");
            this.imageList1.Images.SetKeyName(1, "PartIcon.png");
            // 
            // getRefsWorker
            // 
            this.getRefsWorker.WorkerReportsProgress = true;
            this.getRefsWorker.WorkerSupportsCancellation = true;
            this.getRefsWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.getRefsWorker_DoWork);
            this.getRefsWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.getRefsWorker_ProgressChanged);
            this.getRefsWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.getRefsWorker_RunWorkerCompleted);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar1,
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2});
            this.statusStrip1.Location = new System.Drawing.Point(0, 540);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(465, 22);
            this.statusStrip1.TabIndex = 17;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(0, 17);
            // 
            // CheckoutFilesWorker
            // 
            this.CheckoutFilesWorker.WorkerReportsProgress = true;
            this.CheckoutFilesWorker.WorkerSupportsCancellation = true;
            this.CheckoutFilesWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.CheckoutFilesWorker_DoWork);
            this.CheckoutFilesWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.CheckoutFilesWorker_ProgressChanged);
            this.CheckoutFilesWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.CheckoutFilesWorker_RunWorkerCompleted);
            // 
            // CheckinFilesWorker
            // 
            this.CheckinFilesWorker.WorkerReportsProgress = true;
            this.CheckinFilesWorker.WorkerSupportsCancellation = true;
            this.CheckinFilesWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.CheckinFilesWorker_DoWork);
            this.CheckinFilesWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.CheckinFilesWorker_ProgressChanged);
            this.CheckinFilesWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.CheckinFilesWorker_RunWorkerCompleted);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(210, 32);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(240, 65);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.pictureBox1.TabIndex = 18;
            this.pictureBox1.TabStop = false;
            // 
            // filesListView
            // 
            this.filesListView.CheckBoxes = true;
            this.filesListView.Location = new System.Drawing.Point(12, 166);
            this.filesListView.Name = "filesListView";
            this.filesListView.Size = new System.Drawing.Size(441, 324);
            this.filesListView.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.filesListView.TabIndex = 21;
            this.filesListView.UseCompatibleStateImageBehavior = false;
            this.filesListView.View = System.Windows.Forms.View.Details;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.filesToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(465, 24);
            this.menuStrip1.TabIndex = 22;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // filesToolStripMenuItem
            // 
            this.filesToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.exitToolStripMenuItem});
            this.filesToolStripMenuItem.Name = "filesToolStripMenuItem";
            this.filesToolStripMenuItem.Size = new System.Drawing.Size(42, 20);
            this.filesToolStripMenuItem.Text = "Files";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(103, 22);
            this.openToolStripMenuItem.Text = "Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(103, 22);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "Help";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.aboutToolStripMenuItem.Text = "About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // NPI_Release
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(465, 562);
            this.Controls.Add(this.filesListView);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbBatchRef);
            this.Controls.Add(this.RevisionDescription);
            this.Controls.Add(this.tbRevisionDescription);
            this.Controls.Add(this.ECONumber);
            this.Controls.Add(this.tbECONumber);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Controls.Add(this.VaultsComboBox);
            this.Controls.Add(this.VaultsLabel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "NPI_Release";
            this.Text = "NPI Release";
            this.Load += new System.EventHandler(this.NPI_Release_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Label VaultsLabel;
        internal System.Windows.Forms.ComboBox VaultsComboBox;
        internal System.Windows.Forms.OpenFileDialog OpenFileDialog;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAccept;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TextBox tbECONumber;
        private System.Windows.Forms.Label ECONumber;
        private System.Windows.Forms.TextBox tbRevisionDescription;
        private System.Windows.Forms.Label RevisionDescription;
        private System.Windows.Forms.TextBox tbBatchRef;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ImageList imageList1;
        private System.ComponentModel.BackgroundWorker getRefsWorker;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.ComponentModel.BackgroundWorker CheckoutFilesWorker;
        private System.ComponentModel.BackgroundWorker CheckinFilesWorker;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.ListView filesListView;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem filesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
    }
}

