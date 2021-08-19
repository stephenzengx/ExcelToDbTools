
namespace ExcelTools
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.TxtDirPath = new System.Windows.Forms.TextBox();
            this.BtnStartImport = new System.Windows.Forms.Button();
            this.LbDirPath = new System.Windows.Forms.Label();
            this.BtnValidFile = new System.Windows.Forms.Button();
            this.PanelExcelName = new System.Windows.Forms.Panel();
            this.LbMaxErrorCount = new System.Windows.Forms.Label();
            this.TxtMaxErrorCount = new System.Windows.Forms.TextBox();
            this.LbRemark = new System.Windows.Forms.Label();
            this.LbOrgName = new System.Windows.Forms.Label();
            this.DrpDwnOrgs = new System.Windows.Forms.ComboBox();
            this.BtnReloadJson = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // TxtDirPath
            // 
            this.TxtDirPath.Location = new System.Drawing.Point(288, 155);
            this.TxtDirPath.Name = "TxtDirPath";
            this.TxtDirPath.Size = new System.Drawing.Size(162, 23);
            this.TxtDirPath.TabIndex = 0;
            this.TxtDirPath.Text = "F:/Excel导入测试_816";
            // 
            // BtnStartImport
            // 
            this.BtnStartImport.Location = new System.Drawing.Point(389, 340);
            this.BtnStartImport.Name = "BtnStartImport";
            this.BtnStartImport.Size = new System.Drawing.Size(75, 23);
            this.BtnStartImport.TabIndex = 1;
            this.BtnStartImport.Text = "开始导入";
            this.BtnStartImport.UseVisualStyleBackColor = true;
            this.BtnStartImport.Click += new System.EventHandler(this.ImportBtn_Click);
            // 
            // LbDirPath
            // 
            this.LbDirPath.AutoSize = true;
            this.LbDirPath.Location = new System.Drawing.Point(200, 158);
            this.LbDirPath.Name = "LbDirPath";
            this.LbDirPath.Size = new System.Drawing.Size(80, 17);
            this.LbDirPath.TabIndex = 2;
            this.LbDirPath.Text = "文件夹路径：";
            // 
            // BtnValidFile
            // 
            this.BtnValidFile.Location = new System.Drawing.Point(456, 155);
            this.BtnValidFile.Name = "BtnValidFile";
            this.BtnValidFile.Size = new System.Drawing.Size(91, 23);
            this.BtnValidFile.TabIndex = 6;
            this.BtnValidFile.Text = "开始校验文件";
            this.BtnValidFile.UseVisualStyleBackColor = true;
            this.BtnValidFile.Click += new System.EventHandler(this.BtnValidate_Click);
            // 
            // PanelExcelName
            // 
            this.PanelExcelName.Location = new System.Drawing.Point(200, 184);
            this.PanelExcelName.Name = "PanelExcelName";
            this.PanelExcelName.Size = new System.Drawing.Size(345, 48);
            this.PanelExcelName.TabIndex = 7;
            // 
            // LbMaxErrorCount
            // 
            this.LbMaxErrorCount.AutoSize = true;
            this.LbMaxErrorCount.Location = new System.Drawing.Point(200, 259);
            this.LbMaxErrorCount.Name = "LbMaxErrorCount";
            this.LbMaxErrorCount.Size = new System.Drawing.Size(92, 17);
            this.LbMaxErrorCount.TabIndex = 8;
            this.LbMaxErrorCount.Text = "错误记录阈值：";
            // 
            // TxtMaxErrorCount
            // 
            this.TxtMaxErrorCount.Location = new System.Drawing.Point(288, 256);
            this.TxtMaxErrorCount.Name = "TxtMaxErrorCount";
            this.TxtMaxErrorCount.Size = new System.Drawing.Size(162, 23);
            this.TxtMaxErrorCount.TabIndex = 9;
            this.TxtMaxErrorCount.Text = "100";
            // 
            // LbRemark
            // 
            this.LbRemark.AutoSize = true;
            this.LbRemark.Location = new System.Drawing.Point(200, 296);
            this.LbRemark.Name = "LbRemark";
            this.LbRemark.Size = new System.Drawing.Size(359, 17);
            this.LbRemark.TabIndex = 10;
            this.LbRemark.Text = "备注：导入时，如果检查出的错误记录条数大于阈值,直接退出程序";
            // 
            // LbOrgName
            // 
            this.LbOrgName.AutoSize = true;
            this.LbOrgName.Location = new System.Drawing.Point(190, 112);
            this.LbOrgName.Name = "LbOrgName";
            this.LbOrgName.Size = new System.Drawing.Size(92, 17);
            this.LbOrgName.TabIndex = 13;
            this.LbOrgName.Text = "选择机构名称：";
            // 
            // DrpDwnOrgs
            // 
            this.DrpDwnOrgs.FormattingEnabled = true;
            this.DrpDwnOrgs.Location = new System.Drawing.Point(288, 110);
            this.DrpDwnOrgs.Name = "DrpDwnOrgs";
            this.DrpDwnOrgs.Size = new System.Drawing.Size(259, 25);
            this.DrpDwnOrgs.TabIndex = 19;
            this.DrpDwnOrgs.Text = "请选择机构";
            // 
            // BtnReloadJson
            // 
            this.BtnReloadJson.Location = new System.Drawing.Point(286, 340);
            this.BtnReloadJson.Name = "BtnReloadJson";
            this.BtnReloadJson.Size = new System.Drawing.Size(75, 23);
            this.BtnReloadJson.TabIndex = 20;
            this.BtnReloadJson.Text = "重载配置";
            this.BtnReloadJson.UseVisualStyleBackColor = true;
            this.BtnReloadJson.Click += new System.EventHandler(this.BtnReloadJson_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(748, 515);
            this.Controls.Add(this.BtnReloadJson);
            this.Controls.Add(this.DrpDwnOrgs);
            this.Controls.Add(this.LbOrgName);
            this.Controls.Add(this.LbRemark);
            this.Controls.Add(this.TxtMaxErrorCount);
            this.Controls.Add(this.LbMaxErrorCount);
            this.Controls.Add(this.PanelExcelName);
            this.Controls.Add(this.BtnValidFile);
            this.Controls.Add(this.LbDirPath);
            this.Controls.Add(this.BtnStartImport);
            this.Controls.Add(this.TxtDirPath);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TxtDirPath;
        private System.Windows.Forms.Button BtnStartImport;
        private System.Windows.Forms.Label LbDirPath;
        private System.Windows.Forms.Button BtnValidFile;
        private System.Windows.Forms.Panel PanelExcelName;
        private System.Windows.Forms.Label LbMaxErrorCount;
        private System.Windows.Forms.TextBox TxtMaxErrorCount;
        private System.Windows.Forms.Label LbRemark;
        private System.Windows.Forms.Label LbOrgName;
        private System.Windows.Forms.ComboBox DrpDwnOrgs;
        private System.Windows.Forms.Button BtnReloadJson;
    }
}

