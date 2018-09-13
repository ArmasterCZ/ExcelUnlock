namespace ExcelUnlock
{
    partial class fMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(fMain));
            this.lPath = new System.Windows.Forms.Label();
            this.tbPath = new System.Windows.Forms.TextBox();
            this.bUnlock = new System.Windows.Forms.Button();
            this.lProgress = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lPath
            // 
            this.lPath.AutoSize = true;
            this.lPath.Location = new System.Drawing.Point(13, 13);
            this.lPath.Name = "lPath";
            this.lPath.Size = new System.Drawing.Size(77, 13);
            this.lPath.TabIndex = 0;
            this.lPath.Text = Properties.Resources.FMExcelPath;
            // 
            // tbPath
            // 
            this.tbPath.AllowDrop = true;
            this.tbPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbPath.ForeColor = System.Drawing.Color.DarkGray;
            this.tbPath.Location = new System.Drawing.Point(96, 10);
            this.tbPath.Name = "tbPath";
            this.tbPath.Size = new System.Drawing.Size(347, 20);
            this.tbPath.TabIndex = 1;
            this.tbPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.fMain_DragDrop);
            this.tbPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.fMain_DragEnter);
            this.tbPath.Enter += new System.EventHandler(this.tbPath_GhostTextEnter);
            this.tbPath.Leave += new System.EventHandler(this.tbPath_GhostTextLeave);
            // 
            // bUnlock
            // 
            this.bUnlock.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.bUnlock.Location = new System.Drawing.Point(368, 36);
            this.bUnlock.Name = "bUnlock";
            this.bUnlock.Size = new System.Drawing.Size(75, 23);
            this.bUnlock.TabIndex = 2;
            this.bUnlock.Text = Properties.Resources.FMUnlock;
            this.bUnlock.UseVisualStyleBackColor = true;
            this.bUnlock.Click += new System.EventHandler(this.bUnlock_Click);
            // 
            // lProgress
            // 
            this.lProgress.AutoSize = true;
            this.lProgress.Location = new System.Drawing.Point(12, 41);
            this.lProgress.Name = "lProgress";
            this.lProgress.Size = new System.Drawing.Size(0, 13);
            this.lProgress.TabIndex = 0;
            // 
            // fMain
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(455, 68);
            this.Controls.Add(this.bUnlock);
            this.Controls.Add(this.tbPath);
            this.Controls.Add(this.lProgress);
            this.Controls.Add(this.lPath);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "fMain";
            this.Text = Properties.Resources.FMTitle;
            this.DragDrop += new System.Windows.Forms.DragEventHandler(this.fMain_DragDrop);
            this.DragEnter += new System.Windows.Forms.DragEventHandler(this.fMain_DragEnter);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lPath;
        private System.Windows.Forms.TextBox tbPath;
        private System.Windows.Forms.Button bUnlock;
        private System.Windows.Forms.Label lProgress;
    }
}

