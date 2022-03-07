namespace EveryTeacher
{
    partial class ImportPath
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImportPath));
            this.importOrgPath_btn = new System.Windows.Forms.Button();
            this.text = new System.Windows.Forms.Label();
            this.importOrgPath_txtbx = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.importTchPath_txtbx = new System.Windows.Forms.TextBox();
            this.importTchPath_btn = new System.Windows.Forms.Button();
            this.next_btn = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.exportPath_txtbx = new System.Windows.Forms.TextBox();
            this.exportPath_btn = new System.Windows.Forms.Button();
            this.ckOrg_txt = new System.Windows.Forms.Label();
            this.ckTch_txt = new System.Windows.Forms.Label();
            this.version_txt = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.header_combox = new System.Windows.Forms.ComboBox();
            this.need_mail_cbx = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.sendMail_combox = new System.Windows.Forms.ComboBox();
            this.sendTo_combox = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // importOrgPath_btn
            // 
            resources.ApplyResources(this.importOrgPath_btn, "importOrgPath_btn");
            this.importOrgPath_btn.Name = "importOrgPath_btn";
            this.importOrgPath_btn.UseVisualStyleBackColor = true;
            this.importOrgPath_btn.Click += new System.EventHandler(this.importOrgPath_btn_Click);
            // 
            // text
            // 
            resources.ApplyResources(this.text, "text");
            this.text.Name = "text";
            // 
            // importOrgPath_txtbx
            // 
            resources.ApplyResources(this.importOrgPath_txtbx, "importOrgPath_txtbx");
            this.importOrgPath_txtbx.Name = "importOrgPath_txtbx";
            this.importOrgPath_txtbx.TextChanged += new System.EventHandler(this.ImportPathTextChanged);
            // 
            // label10
            // 
            resources.ApplyResources(this.label10, "label10");
            this.label10.Name = "label10";
            // 
            // importTchPath_txtbx
            // 
            resources.ApplyResources(this.importTchPath_txtbx, "importTchPath_txtbx");
            this.importTchPath_txtbx.Name = "importTchPath_txtbx";
            // 
            // importTchPath_btn
            // 
            resources.ApplyResources(this.importTchPath_btn, "importTchPath_btn");
            this.importTchPath_btn.Name = "importTchPath_btn";
            this.importTchPath_btn.UseVisualStyleBackColor = true;
            this.importTchPath_btn.Click += new System.EventHandler(this.importTchPath_btn_Click);
            // 
            // next_btn
            // 
            resources.ApplyResources(this.next_btn, "next_btn");
            this.next_btn.Name = "next_btn";
            this.next_btn.UseVisualStyleBackColor = true;
            this.next_btn.Click += new System.EventHandler(this.next_btn_Click);
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            // 
            // exportPath_txtbx
            // 
            resources.ApplyResources(this.exportPath_txtbx, "exportPath_txtbx");
            this.exportPath_txtbx.Name = "exportPath_txtbx";
            // 
            // exportPath_btn
            // 
            resources.ApplyResources(this.exportPath_btn, "exportPath_btn");
            this.exportPath_btn.Name = "exportPath_btn";
            this.exportPath_btn.UseVisualStyleBackColor = true;
            this.exportPath_btn.Click += new System.EventHandler(this.exportPath_btn_Click);
            // 
            // ckOrg_txt
            // 
            resources.ApplyResources(this.ckOrg_txt, "ckOrg_txt");
            this.ckOrg_txt.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.ckOrg_txt.Name = "ckOrg_txt";
            // 
            // ckTch_txt
            // 
            resources.ApplyResources(this.ckTch_txt, "ckTch_txt");
            this.ckTch_txt.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(0)))));
            this.ckTch_txt.Name = "ckTch_txt";
            // 
            // version_txt
            // 
            resources.ApplyResources(this.version_txt, "version_txt");
            this.version_txt.Name = "version_txt";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // header_combox
            // 
            resources.ApplyResources(this.header_combox, "header_combox");
            this.header_combox.FormattingEnabled = true;
            this.header_combox.Name = "header_combox";
            // 
            // need_mail_cbx
            // 
            resources.ApplyResources(this.need_mail_cbx, "need_mail_cbx");
            this.need_mail_cbx.Name = "need_mail_cbx";
            this.need_mail_cbx.UseVisualStyleBackColor = true;
            this.need_mail_cbx.CheckedChanged += new System.EventHandler(this.needMailChanged);
            // 
            // label3
            // 
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // label5
            // 
            resources.ApplyResources(this.label5, "label5");
            this.label5.Name = "label5";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // label6
            // 
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            resources.ApplyResources(this.label6, "label6");
            this.label6.Name = "label6";
            // 
            // sendMail_combox
            // 
            resources.ApplyResources(this.sendMail_combox, "sendMail_combox");
            this.sendMail_combox.FormattingEnabled = true;
            this.sendMail_combox.Name = "sendMail_combox";
            // 
            // sendTo_combox
            // 
            resources.ApplyResources(this.sendTo_combox, "sendTo_combox");
            this.sendTo_combox.FormattingEnabled = true;
            this.sendTo_combox.Name = "sendTo_combox";
            // 
            // label7
            // 
            resources.ApplyResources(this.label7, "label7");
            this.label7.Name = "label7";
            // 
            // label8
            // 
            resources.ApplyResources(this.label8, "label8");
            this.label8.Name = "label8";
            // 
            // ImportPath
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.sendMail_combox);
            this.Controls.Add(this.sendTo_combox);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.need_mail_cbx);
            this.Controls.Add(this.header_combox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.version_txt);
            this.Controls.Add(this.ckTch_txt);
            this.Controls.Add(this.ckOrg_txt);
            this.Controls.Add(this.exportPath_btn);
            this.Controls.Add(this.exportPath_txtbx);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.next_btn);
            this.Controls.Add(this.importTchPath_btn);
            this.Controls.Add(this.importTchPath_txtbx);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.importOrgPath_txtbx);
            this.Controls.Add(this.text);
            this.Controls.Add(this.importOrgPath_btn);
            this.Name = "ImportPath";
            this.Load += new System.EventHandler(this.ImportPath_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button importOrgPath_btn;
        private System.Windows.Forms.Label text;
        private System.Windows.Forms.TextBox importOrgPath_txtbx;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox importTchPath_txtbx;
        private System.Windows.Forms.Button importTchPath_btn;
        private System.Windows.Forms.Button next_btn;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox exportPath_txtbx;
        private System.Windows.Forms.Button exportPath_btn;
        private System.Windows.Forms.Label ckOrg_txt;
        private System.Windows.Forms.Label ckTch_txt;
        private System.Windows.Forms.Label version_txt;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox header_combox;
        private System.Windows.Forms.CheckBox need_mail_cbx;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox sendMail_combox;
        private System.Windows.Forms.ComboBox sendTo_combox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
    }
}

