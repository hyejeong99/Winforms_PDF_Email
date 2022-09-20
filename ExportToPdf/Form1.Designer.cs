namespace ExportToPdf
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btn_savePDF = new System.Windows.Forms.Button();
            this.btn_printPDF = new System.Windows.Forms.Button();
            this.btn_sendEmail = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(25, 22);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(763, 320);
            this.dataGridView1.TabIndex = 0;
            // 
            // btn_savePDF
            // 
            this.btn_savePDF.Font = new System.Drawing.Font("굴림", 12F, System.Drawing.FontStyle.Bold);
            this.btn_savePDF.Location = new System.Drawing.Point(190, 382);
            this.btn_savePDF.Name = "btn_savePDF";
            this.btn_savePDF.Size = new System.Drawing.Size(131, 33);
            this.btn_savePDF.TabIndex = 1;
            this.btn_savePDF.Text = "PDF로 저장";
            this.btn_savePDF.UseVisualStyleBackColor = true;
            this.btn_savePDF.Click += new System.EventHandler(this.button1_Click);
            // 
            // btn_printPDF
            // 
            this.btn_printPDF.Font = new System.Drawing.Font("굴림", 12F, System.Drawing.FontStyle.Bold);
            this.btn_printPDF.Location = new System.Drawing.Point(327, 382);
            this.btn_printPDF.Name = "btn_printPDF";
            this.btn_printPDF.Size = new System.Drawing.Size(131, 33);
            this.btn_printPDF.TabIndex = 2;
            this.btn_printPDF.Text = "PDF로 인쇄";
            this.btn_printPDF.UseVisualStyleBackColor = true;
            // 
            // btn_sendEmail
            // 
            this.btn_sendEmail.Font = new System.Drawing.Font("굴림", 12F, System.Drawing.FontStyle.Bold);
            this.btn_sendEmail.Location = new System.Drawing.Point(464, 382);
            this.btn_sendEmail.Name = "btn_sendEmail";
            this.btn_sendEmail.Size = new System.Drawing.Size(131, 33);
            this.btn_sendEmail.TabIndex = 3;
            this.btn_sendEmail.Text = "이메일 전송";
            this.btn_sendEmail.UseVisualStyleBackColor = true;
            this.btn_sendEmail.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btn_sendEmail);
            this.Controls.Add(this.btn_printPDF);
            this.Controls.Add(this.btn_savePDF);
            this.Controls.Add(this.dataGridView1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btn_savePDF;
        private System.Windows.Forms.Button btn_printPDF;
        private System.Windows.Forms.Button btn_sendEmail;
    }
}

