namespace EveryTeacher
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_select_file = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.lb_sys_info = new System.Windows.Forms.Label();
            this.lb_load_data_count = new System.Windows.Forms.Label();
            this.btn_select_sheet = new System.Windows.Forms.Button();
            this.text = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btn_select_file
            // 
            this.btn_select_file.Location = new System.Drawing.Point(12, 22);
            this.btn_select_file.Name = "btn_select_file";
            this.btn_select_file.Size = new System.Drawing.Size(89, 29);
            this.btn_select_file.TabIndex = 0;
            this.btn_select_file.Text = "按下";
            this.btn_select_file.UseVisualStyleBackColor = true;
            this.btn_select_file.Click += new System.EventHandler(this.btn_select_file_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(268, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(520, 426);
            this.dataGridView1.TabIndex = 1;
            // 
            // lb_sys_info
            // 
            this.lb_sys_info.AutoSize = true;
            this.lb_sys_info.Location = new System.Drawing.Point(13, 238);
            this.lb_sys_info.Name = "lb_sys_info";
            this.lb_sys_info.Size = new System.Drawing.Size(41, 15);
            this.lb_sys_info.TabIndex = 4;
            this.lb_sys_info.Text = "label1";
            // 
            // lb_load_data_count
            // 
            this.lb_load_data_count.AutoSize = true;
            this.lb_load_data_count.Location = new System.Drawing.Point(16, 287);
            this.lb_load_data_count.Name = "lb_load_data_count";
            this.lb_load_data_count.Size = new System.Drawing.Size(41, 15);
            this.lb_load_data_count.TabIndex = 5;
            this.lb_load_data_count.Text = "label1";
            // 
            // btn_select_sheet
            // 
            this.btn_select_sheet.Location = new System.Drawing.Point(16, 78);
            this.btn_select_sheet.Name = "btn_select_sheet";
            this.btn_select_sheet.Size = new System.Drawing.Size(75, 23);
            this.btn_select_sheet.TabIndex = 6;
            this.btn_select_sheet.Text = "button1";
            this.btn_select_sheet.UseVisualStyleBackColor = true;
            this.btn_select_sheet.Click += new System.EventHandler(this.btn_select_sheet_Click);
            // 
            // text
            // 
            this.text.AutoSize = true;
            this.text.Location = new System.Drawing.Point(16, 334);
            this.text.Name = "text";
            this.text.Size = new System.Drawing.Size(41, 15);
            this.text.TabIndex = 7;
            this.text.Text = "label1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.text);
            this.Controls.Add(this.btn_select_sheet);
            this.Controls.Add(this.lb_load_data_count);
            this.Controls.Add(this.lb_sys_info);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btn_select_file);
            this.Name = "Form1";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_select_file;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label lb_sys_info;
        private System.Windows.Forms.Label lb_load_data_count;
        private System.Windows.Forms.Button btn_select_sheet;
        private System.Windows.Forms.Label text;
    }
}

