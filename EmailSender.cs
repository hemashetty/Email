using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BobMISTransmit
{
    public partial class EmailSender : Form
    {
        public EmailSender()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            EmailSender emailSender = new EmailSender();
            emailSender.SendEmailWithExcelInBody(@"C:\Users\shett\Downloads\Final_Net_Gross_Sales_09.07.2024.xslx");
            this.Close();
        }
        public void SendEmailWithExcelInBody(string excelFilePath)
        {
            string smtpAddress = "smtp.example.com"; // Your SMTP server
            int portNumber = 587; // Your SMTP port
            bool enableSSL = true;

            string emailFrom = "your_email@example.com";
            string password = "your_password";
            string emailTo = "recipient@example.com";
            string subject = "MIS Report";

            // Convert Excel to HTML table
            string body = ConvertExcelToHtmlTable(excelFilePath);

            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress(emailFrom);
                mail.To.Add(emailTo);
                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = true; // Email body is HTML

                using (SmtpClient smtp = new SmtpClient(smtpAddress, portNumber))
                {
                    smtp.Credentials = new NetworkCredential(emailFrom, password);
                    smtp.EnableSsl = enableSSL;
                    smtp.Send(mail);
                }
            }
        }

        private string ConvertExcelToHtmlTable(string excelFilePath)
        {
            // Assuming you're using a library like ClosedXML to read Excel files
            // Read Excel data into a DataTable
            DataTable dt = new DataTable();

            // Load your DataTable with Excel data here

            // Convert DataTable to HTML
            StringBuilder html = new StringBuilder();
            html.Append("<table border='1'>");

            // Add table headers
            html.Append("<tr>");
            foreach (DataColumn column in dt.Columns)
            {
                html.Append("<th>");
                html.Append(column.ColumnName);
                html.Append("</th>");
            }
            html.Append("</tr>");

            // Add table rows
            foreach (DataRow row in dt.Rows)
            {
                html.Append("<tr>");
                foreach (var cell in row.ItemArray)
                {
                    html.Append("<td>");
                    html.Append(cell.ToString());
                    html.Append("</td>");
                }
                html.Append("</tr>");
            }

            html.Append("</table>");
            return html.ToString();
        }
    }
}
