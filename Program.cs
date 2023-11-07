using Microsoft.Extensions.Configuration;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace SendMailOutlook
{
    internal class Program
    {
        static void Main(string[] args)
        {
            IConfiguration configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json")
                .Build();

            var smtpSettings = configuration.GetSection("SmtpSettings");
            var emailSettings = configuration.GetSection("EmailSettings");
            var folderPath = configuration["FolderPath"];

            MailMessage message = new MailMessage();

            try
            {
                SmtpClient client = new SmtpClient(smtpSettings["SmtpServer"])
                {
                    Port = int.TryParse(smtpSettings["Port"], out int port) ? port : 587,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(smtpSettings["Username"], smtpSettings["Password"]),
                    EnableSsl = true
                };

                string fromEmail = emailSettings["FromEmail"]?.ToString() ?? throw new InvalidOperationException("FromEmail 不可為空值");
                string toEmails = emailSettings["ToEmail"]?.ToString() ?? throw new InvalidOperationException("ToEmail 不可為空值");
                string[] recipients = toEmails.Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var recipient in recipients)
                {
                    message.To.Add(recipient.Trim());
                }

                bool validChoice = false;

                while (!validChoice)
                {
                    Console.WriteLine("是否要傳送附加檔案（Y/N）？:");
                    char choice = Console.ReadKey().KeyChar;
                    Console.WriteLine();

                    if (choice == 'y' || choice == 'Y')
                    {
                        // 獲取資料夾內的所有檔案
                        string[] filePaths = Directory.GetFiles(folderPath);

                        foreach (string filePath in filePaths)
                        {
                            if (File.Exists(filePath))
                            {
                                Attachment attachment = new Attachment(filePath);
                                message.Attachments.Add(attachment);
                            }
                            else
                            {
                                Console.WriteLine($"檔案 {filePath} 不存在，將不會加入附件。");
                            }
                        }

                        Console.WriteLine("傳送檔案完成");
                        validChoice = true;
                    }
                    else if (choice == 'n' || choice == 'N')
                    {
                        Console.WriteLine("不傳送檔案");
                        validChoice = true;
                    }
                    else
                    {
                        Console.WriteLine("無效的選擇，請重新輸入。");
                    }
                }

                // 設定郵件主旨和內容
                message.From = new MailAddress(fromEmail);
                message.Subject = "test mail";
                message.Body = "<h1>測試</h1>";
                message.IsBodyHtml = true;

                // 發送郵件
                client.Send(message);
                Console.WriteLine("郵件發送成功！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"郵件發送失敗：{ex.Message}");
            }
            finally
            {
                // 釋放附件資源
                foreach (var attachment in message.Attachments)
                {
                    attachment.Dispose();
                }
            }
        }
    }
}
