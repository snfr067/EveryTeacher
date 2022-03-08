namespace EveryTeacher
{
    partial class SplitExcel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SplitExcel));
            this.tchFile_pbar = new System.Windows.Forms.ProgressBar();
            this.Over_btn = new System.Windows.Forms.Button();
            this.tchFileP_txt = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // tchFile_pbar
            // 
            this.tchFile_pbar.Location = new System.Drawing.Point(31, 28);
            this.tchFile_pbar.Name = "tchFile_pbar";
            this.tchFile_pbar.Size = new System.Drawing.Size(580, 27);
            this.tchFile_pbar.TabIndex = 17;
            // 
            // Over_btn
            // 
            this.Over_btn.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Over_btn.Location = new System.Drawing.Point(486, 73);
            this.Over_btn.Name = "Over_btn";
            this.Over_btn.Size = new System.Drawing.Size(125, 35);
            this.Over_btn.TabIndex = 19;
            this.Over_btn.Text = "取消";
            this.Over_btn.UseVisualStyleBackColor = true;
            this.Over_btn.Click += new System.EventHandler(this.Over_btnClick);
            // 
            // tchFileP_txt
            // 
            this.tchFileP_txt.AutoSize = true;
            this.tchFileP_txt.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.tchFileP_txt.Location = new System.Drawing.Point(27, 80);
            this.tchFileP_txt.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.tchFileP_txt.Name = "tchFileP_txt";
            this.tchFileP_txt.Size = new System.Drawing.Size(69, 20);
            this.tchFileP_txt.TabIndex = 16;
            this.tchFileP_txt.Text = "計算中...";
            // 
            // SplitExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(641, 128);
            this.Controls.Add(this.Over_btn);
            this.Controls.Add(this.tchFile_pbar);
            this.Controls.Add(this.tchFileP_txt);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(500, 500);
            this.Name = "SplitExcel";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "檔案產出進度";
            this.Load += new System.EventHandler(this.SplitExcelLoad);
            this.Shown += new System.EventHandler(this.SplitExcelShown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ProgressBar tchFile_pbar;
        private System.Windows.Forms.Button Over_btn;
        private System.Windows.Forms.Label tchFileP_txt;
    }
}