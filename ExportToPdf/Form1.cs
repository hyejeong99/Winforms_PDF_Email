using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Net.Mail;

namespace ExportToPdf
{
    public partial class Form1 : Form
    {
        public string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\82104\source\repos\ExportToPdf\ExportToPdf\Database1.mdf;Integrated Security=True;Connect Timeout=30";
        public Form1()
        {
            InitializeComponent();

            //데이터베이스 값 가져와서 보여주기
            try
            {
                /*string query = @"SELECT * FROM tblEmployee";
                DataSet ds = new System.Data.DataSet();
                OracleDBManager.Instance.ExecuteDsQuery(ds, query);*/
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = @"SELECT * FROM tblEmployee";
                SqlDataReader sdr = cmd.ExecuteReader();

                /*if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                    return;
                dataGridView1.DataSource = ds.Tables[0];*/

                while (sdr.Read() == true)
                {
                    dataGridView1.DataSource = sdr["EmployeeId"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SqlConnection sqlCon;
            string conString = null;
            string sqlQuery = null;

            conString = "Data Source=.;Initial Catalog=DemoTest;Integrated Security=SSPI;";
            //conString= @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\82104\source\repos\lightP03\lightP03\Database1.mdf;Integrated Security=True;Connect Timeout=30";
            //sqlCon = new SqlConnection(conString);
            sqlCon = new SqlConnection(connectionString);

            sqlCon.Open();
            sqlQuery = "SELECT * FROM tblEmployee";
            SqlDataAdapter dscmd = new SqlDataAdapter(sqlQuery, sqlCon);
            DataTable dtData = new DataTable();
            dscmd.Fill(dtData);
            dataGridView1.DataSource = dtData;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "PDF (*.pdf)|*.pdf";
                sfd.FileName = "Output.pdf";
                bool fileError = false;
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(sfd.FileName))
                    {
                        try
                        {
                            File.Delete(sfd.FileName);
                        }
                        catch (IOException ex)
                        {
                            fileError = true;
                            MessageBox.Show("It wasn't possible to write the data to the disk." + ex.Message);
                        }
                    }
                    if (!fileError)
                    {
                        try
                        {
                            PdfPTable pdfTable = new PdfPTable(dataGridView1.Columns.Count);
                            pdfTable.DefaultCell.Padding = 3;
                            pdfTable.WidthPercentage = 100;
                            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT;

                            foreach (DataGridViewColumn column in dataGridView1.Columns)
                            {
                                PdfPCell cell = new PdfPCell(new Phrase(column.HeaderText));
                                pdfTable.AddCell(cell);
                            }

                            foreach (DataGridViewRow row in dataGridView1.Rows)
                            {
                                foreach (DataGridViewCell cell in row.Cells)
                                {
                                    pdfTable.AddCell(cell.Value.ToString());
                                }
                            }

                            using (FileStream stream = new FileStream(sfd.FileName, FileMode.Create))
                            {
                                Document pdfDoc = new Document(PageSize.A4, 10f, 20f, 20f, 10f);
                                PdfWriter.GetInstance(pdfDoc, stream);
                                pdfDoc.Open();
                                pdfDoc.Add(pdfTable);
                                pdfDoc.Close();
                                stream.Close();
                            }

                            MessageBox.Show("Data Exported Successfully !!!", "Info");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error :" + ex.Message);
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No Record To Export !!!", "Info");
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            MailMessage mail = new MailMessage();
            try
            {   // 보내는 사람 메일, 이름, 인코딩(UTF-8)      
                mail.From = new MailAddress("gg2060gg@gmail.com", "명월일지", System.Text.Encoding.UTF8);
                // 받는 사람 메일      
                mail.To.Add("gg2060gg@gmail.com");
                // 참조 사람 메일      
                mail.CC.Add("gg2060gg@gmail.com");
                // 비공개 참조 사람 메일      
                mail.Bcc.Add("gg2060gg@gmail.com");
                // 메일 제목      
                mail.Subject = "PDF 파일";
                // 본문 내용      
                mail.Body = "<html><body>hello wrold</body></html>";
                // 본문 내용 포멧의 타입 (true의 경우 Html 포멧으로)      
                mail.IsBodyHtml = true;

                // 메일 제목과 본문의 인코딩 타입(UTF-8)      
                mail.SubjectEncoding = System.Text.Encoding.UTF8;
                mail.BodyEncoding = System.Text.Encoding.UTF8;
                // 첨부 파일 (Stream과 파일 이름)      
                mail.Attachments.Add(new Attachment(new FileStream(@"C:\Users\82104\source\repos\ExportToPdf\test1.pdf", FileMode.Open, FileAccess.Read), "test1.pdf"));
                //mail.Attachments.Add(new Attachment(new FileStream(@"D:\test2.zip", FileMode.Open, FileAccess.Read), "test2.zip"));

                // smtp 서버 주소      
                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                // smtp 포트      
                SmtpServer.Port = 587;
                // smtp 인증      
                SmtpServer.Credentials = new System.Net.NetworkCredential("gg2060gg", "vwsfdgymtyfqqeuu");
                // SSL 사용 여부      
                SmtpServer.EnableSsl = true;
                // 발송      
                SmtpServer.Send(mail);
            }
            finally
            {
                // 첨부 파일 Stream 닫기      
                foreach (var attach in mail.Attachments)
                {
                    attach.ContentStream.Close();
                }
            }
        }
    }
}
