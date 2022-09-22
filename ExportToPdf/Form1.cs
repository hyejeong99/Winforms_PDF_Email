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
using System.Collections;
using System.Drawing.Printing;
using Font = System.Drawing.Font;
using Rectangle = System.Drawing.Rectangle;

namespace ExportToPdf
{
    public partial class Form1 : Form
    {
        public string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\82104\source\repos\ExportToPdf\ExportToPdf\Database1.mdf;Integrated Security=True;Connect Timeout=30";
        public Form1()
        {
            InitializeComponent();
            //datagridview 값 추가
            dataGridView1.Rows.Add("1", "2", "3", "4");



            //데이터베이스 값 가져와서 보여주기
            /*try
            {
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;
                cmd.CommandText = @"SELECT * FROM tblEmployee";
                SqlDataReader sdr = cmd.ExecuteReader();

                while (sdr.Read() == true)
                {
                    dataGridView1.DataSource = sdr["EmployeeId"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }*/
        }

    private void Form1_Load(object sender, EventArgs e)
        {
            /*SqlConnection sqlCon;
            string conString = null;
            string sqlQuery = null;

            conString = "Data Source=.;Initial Catalog=DemoTest;Integrated Security=SSPI;";
            sqlCon = new SqlConnection(connectionString);

            sqlCon.Open();
            sqlQuery = "SELECT * FROM tblEmployee";
            SqlDataAdapter dscmd = new SqlDataAdapter(sqlQuery, sqlCon);
            DataTable dtData = new DataTable();
            dscmd.Fill(dtData);
            dataGridView1.DataSource = dtData;*/
        }

        private void button1_Click(object sender, EventArgs e)
        {//PDF로 저장
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
        {//이메일 전송
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
        class ClsPrint
        {
            //프린트 기능 시작
            int iCellHeight = 0;
            //Used to get/set the datagridview cell height    
            int iTotalWidth = 0;
            //           
            int iRow = 0;
            //Used as counter       
            bool bFirstPage = false;
            //Used to check whether we are printing first page        
            bool bNewPage = false;
            // Used to check whether we are printing a new page        
            int iHeaderHeight = 0; //Used for the header height    
            StringFormat strFormat; //Used to format the grid rows.       
            ArrayList arrColumnLefts = new ArrayList();//Used to save left coordinates of columns     
            ArrayList arrColumnWidths = new ArrayList();//Used to save column widths       
            private PrintDocument _printDocument = new PrintDocument();
            private DataGridView gw = new DataGridView();
            private string _ReportHeader;

            public ClsPrint(DataGridView gridview, string ReportHeader)
            {
                _printDocument.PrintPage += new PrintPageEventHandler(_printDocument_PrintPage);
                _printDocument.BeginPrint += new PrintEventHandler(_printDocument_BeginPrint);
                gw = gridview;
                _ReportHeader = ReportHeader;
            }
            public void PrintForm()
            {
                PrintPreviewDialog objPPdialog = new PrintPreviewDialog();
                objPPdialog.Document = _printDocument;
                objPPdialog.ShowDialog();
            }
            private void _printDocument_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
            {
                int iLeftMargin = e.MarginBounds.Left;
                //Set the top margin               
                int iTopMargin = e.MarginBounds.Top;
                //Whether more pages have to print or not             
                bool bMorePagesToPrint = false;
                int iTmpWidth = 0;
                //For the first page to print set the cell width and header height    
                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in gw.Columns)
                    {
                        iTmpWidth = (int)(((Math.Floor((double)((double)GridCol.Width / (double)iTotalWidth * (double)iTotalWidth * ((double)e.MarginBounds.Width / (double)iTotalWidth))))) * 0.66);
                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText, GridCol.InheritedStyle.Font, iTmpWidth).Height) + 11;
                        // Save width and height of headers                     
                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidths.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }
                //Loop till all the grid rows not get printed       
                while (iRow <= gw.Rows.Count - 1)
                {
                    DataGridViewRow GridRow = gw.Rows[iRow];
                    //Set the cell height                 
                    iCellHeight = GridRow.Height + 5;
                    int iCount = 0;
                    //Check whether the current page settings allows more rows to print     
                    if (iTopMargin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            //Draw Header             
                            e.Graphics.DrawString(_ReportHeader,
                                new Font(gw.Font, FontStyle.Bold),
                                Brushes.Black, e.MarginBounds.Left,
                                e.MarginBounds.Top - e.Graphics.MeasureString(_ReportHeader,
                                new Font(gw.Font, FontStyle.Bold),
                                e.MarginBounds.Width).Height - 13);
                            String strDate = "";
                            //Draw Date              
                            e.Graphics.DrawString(strDate,
                                new Font(gw.Font, FontStyle.Bold), Brushes.Black,
                                e.MarginBounds.Left + (e.MarginBounds.Width - e.Graphics.MeasureString(strDate, new Font(gw.Font, FontStyle.Bold), e.MarginBounds.Width).Width),
                                e.MarginBounds.Top - e.Graphics.MeasureString(_ReportHeader, new Font(new Font(gw.Font, FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);
                            //Draw Columns                                      
                            iTopMargin = e.MarginBounds.Top;
                            DataGridViewColumn[] _GridCol = new DataGridViewColumn[gw.Columns.Count];
                            int colcount = 0;                            //Convert ltr to rtl                       
                            foreach (DataGridViewColumn GridCol in gw.Columns)
                            {
                                _GridCol[colcount++] = GridCol;
                            }
                            // for (int i = (_GridCol.Count() - 1); i >= 0; i--)            
                            for (int i = 0; i <= (_GridCol.Count() - 1); i++)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));
                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight));
                                e.Graphics.DrawString(_GridCol[i].HeaderText,
                                    _GridCol[i].InheritedStyle.Font,
                                    new SolidBrush(_GridCol[i].InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMargin,
                                    (int)arrColumnWidths[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMargin += iHeaderHeight;
                        }
                        iCount = 0;
                        DataGridViewCell[] _GridCell = new DataGridViewCell[GridRow.Cells.Count];
                        int cellcount = 0;
                        //Convert ltr to rtl             
                        foreach (DataGridViewCell Cel in GridRow.Cells)
                        {
                            _GridCell[cellcount++] = Cel;
                        }
                        //Draw Columns Contents          
                        //for (int i = (_GridCell.Count() - 1); i >= 0; i--)     
                        for (int i = 0; i <= (_GridCell.Count() - 1); i++)
                        {
                            if (_GridCell[i].Value != null)
                            {
                                e.Graphics.DrawString(_GridCell[i].FormattedValue.ToString(),
                                    _GridCell[i].InheritedStyle.Font,
                                    new SolidBrush(_GridCell[i].InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount],
                                    (float)iTopMargin,
                                    (int)arrColumnWidths[iCount], (float)iCellHeight),
                                    strFormat);
                            }
                            //Drawing Cells Borders                      
                            e.Graphics.DrawRectangle(Pens.Black,
                                new Rectangle((int)arrColumnLefts[iCount], iTopMargin,
                                (int)arrColumnWidths[iCount], iCellHeight));
                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMargin += iCellHeight;
                }
                //If more lines exist, print another page.           
                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
                //}             
                //catch (Exception exc)              
                //{             
                //    MessageBox.Show(exc.Message, "Error", MessageBoxButtons.OK,       
                //       MessageBoxIcon.Error);                //
            }

            private void _printDocument_BeginPrint(object sender, PrintEventArgs e)
            {
                try
                {
                    strFormat = new StringFormat();
                    strFormat.Alignment = StringAlignment.Near;
                    strFormat.LineAlignment = StringAlignment.Center;
                    strFormat.Trimming = StringTrimming.EllipsisCharacter;
                    arrColumnLefts.Clear();
                    arrColumnWidths.Clear();
                    iCellHeight = 0;
                    iRow = 0;
                    bFirstPage = true;
                    bNewPage = true;
                    // Calculating Total Widths           
                    iTotalWidth = 0;
                    foreach (DataGridViewColumn dgvGridCol in gw.Columns)
                    {
                        iTotalWidth += dgvGridCol.Width;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        

        private void btn_printPDF_Click(object sender, EventArgs e)
        {//인쇄 기능
            ClsPrint _ClsPrint = new ClsPrint(dataGridView1, "출력 제목"); _ClsPrint.PrintForm();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}