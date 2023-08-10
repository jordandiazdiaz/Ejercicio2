using System;
using System.Data.SqlClient;
using System.Net.Mail;
using OfficeOpenXml;

namespace EmployeeReportApp
{
    class Program
    {
        static void Main(string[] args)
        {
            string startDateEnv = Environment.GetEnvironmentVariable("START_DATE");
            string endDateEnv = Environment.GetEnvironmentVariable("END_DATE");
            DateTime startDate = DateTime.Parse(startDateEnv ?? "2021-01-01");
            DateTime endDate = DateTime.Parse(endDateEnv ?? "2022-12-31");

            string connString = "Server=localhost;Database=MiBaseDeDatos;Trusted_Connection=True;";
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                string query = @"SELECT * FROM Employees WHERE AdmissionDate BETWEEN @startDate AND @endDate";
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@startDate", startDate);
                    cmd.Parameters.AddWithValue("@endDate", endDate);

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        using (ExcelPackage package = new ExcelPackage())
                        {
                            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Reporte");
                            // Agregar los encabezados aquí
                            int row = 2;
                            while (reader.Read())
                            {
                                // Rellenar el archivo Excel aquí
                                // Ejemplo:
                                worksheet.Cells[row, 1].Value = reader.GetInt32(0); // Id
                                worksheet.Cells[row, 2].Value = reader.GetString(1); // Name
                                // ...
                                row++;
                            }

                            // Guardar y enviar el archivo Excel
                            string path = "report.xlsx";
                            package.SaveAs(new System.IO.FileInfo(path));
                            SendEmail(path);
                        }
                    }
                }
            }
        }

        static void SendEmail(string filePath)
        {
            MailMessage mail = new MailMessage();
            mail.From = new MailAddress("tu_correo@example.com");
            mail.To.Add("franco.paredes@oechsle.pe");
            mail.Subject = "Reporte Empleados - Examen Técnico Oechsle";
            mail.Body = "Aquí tienes el reporte solicitado.";

            Attachment attachment = new Attachment(filePath);
            mail.Attachments.Add(attachment);

            SmtpClient smtpClient = new SmtpClient("smtp.example.com")
            {
                Port = 587,
                Credentials = new System.Net.NetworkCredential("tu_correo@example.com", "tu_contraseña"),
                EnableSsl = true,
            };

            smtpClient.Send(mail);
        }
    }
}
