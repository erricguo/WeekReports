namespace WeekReports
{
    partial class MsgForm
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
            this.timer1 = new System.Windows.Forms.Timer();
            this.me01 = new DevExpress.XtraEditors.MemoEdit();
            this.btnOK = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.me01.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // me01
            // 
            this.me01.Dock = System.Windows.Forms.DockStyle.Top;
            this.me01.EditValue = "測試文字";
            this.me01.Location = new System.Drawing.Point(0, 0);
            this.me01.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.me01.Name = "me01";
            this.me01.Properties.AcceptsReturn = false;
            this.me01.Properties.Appearance.Font = new System.Drawing.Font("YaHei Consolas Hybrid", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.me01.Properties.Appearance.Options.UseFont = true;
            this.me01.Properties.ReadOnly = true;
            this.me01.Properties.ScrollBars = System.Windows.Forms.ScrollBars.None;
            this.me01.Size = new System.Drawing.Size(499, 216);
            this.me01.TabIndex = 0;
            this.me01.EditValueChanged += new System.EventHandler(this.me01_EditValueChanged);
            // 
            // btnOK
            // 
            this.btnOK.Appearance.Font = new System.Drawing.Font("YaHei Consolas Hybrid", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.Appearance.Options.UseFont = true;
            this.btnOK.ImageToTextAlignment = DevExpress.XtraEditors.ImageAlignToText.LeftCenter;
            this.btnOK.Location = new System.Drawing.Point(194, 224);
            this.btnOK.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(111, 39);
            this.btnOK.TabIndex = 20;
            this.btnOK.Text = "確定";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // MsgForm
            // 
            this.Appearance.Options.UseFont = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 21F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(499, 270);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.me01);
            this.Font = new System.Drawing.Font("YaHei Consolas Hybrid", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "MsgForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "MsgForm";
            this.Load += new System.EventHandler(this.MsgForm_Load);
            this.Shown += new System.EventHandler(this.MsgForm_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.me01.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private DevExpress.XtraEditors.SimpleButton btnOK;
        private System.Windows.Forms.Timer timer1;
        private DevExpress.XtraEditors.MemoEdit me01;
    }
}