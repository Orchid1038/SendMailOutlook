using Microsoft.Extensions.Configuration;
using SFT.Core;
using System.Data;
using System.Data.SqlClient;
using System.Net;
using System.Net.Mail;

namespace SendMailOutlook
{
    internal class Program
    {
        private static readonly string LogFolderPath = "E:\\SendMailOutlook\\Logs";
        private static readonly string LogFilePath = $"{LogFolderPath}\\{DateTime.Now:yyyyMMdd-HHmmss}_Log.txt";
        private static Scheduler? scheduler;
        private static readonly string EmailHtmlFilePath = "Email.html";
        private static readonly string EmailSubject = "測試用信件(排程&Excel排版)";

        static void Main(string[] args)
        {
            IConfiguration configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json")
                .Build();

            var smtpSettings = configuration.GetSection("SmtpSettings");
            var emailSettings = configuration.GetSection("EmailSettings");
            string? folderPath = configuration["FolderPath"];


            scheduler = new Scheduler(smtpSettings, emailSettings, folderPath);

            scheduler.Start();
            scheduler.Stop();
        }

        private static readonly object lockObject = new object();
        /// <summary>
        /// 主要是負責推送郵件的執行緒
        /// </summary>
        /// <param name="smtpSettings">預設已經有了，為SMTP的設定</param>
        /// <param name="emailSettings">這是EMAIL的設定，預留使用</param>
        /// <param name="folderPath">路徑，詳情查看JSON檔案</param>
        public static void DoExportAndSendEmail(IConfigurationSection smtpSettings, IConfigurationSection emailSettings, string folderPath)
        {
            lock (lockObject)
            {
                DataSet emailDataSet = EmailQuest();

                // 處理 emailDataSet DataSet
                ProcessDataSet(emailDataSet, folderPath, "emailDataSet", smtpSettings);
            }
        }

        /// <summary>
        /// 判斷要寄送哪一個excel的方法，日後維護使用
        /// </summary>
        /// <param name="dataSet">這裡的dataset傳送的是EmailQuest不用帶參數進去</param>
        /// <param name="folderPath"></param>
        /// <param name="dataSetName">寫log紀錄使用，哪個方法傳送進來的參數比較好判別</param>
        /// <param name="smtpSettings"></param>
        private static void ProcessDataSet(DataSet dataSet, string folderPath, string dataSetName, IConfigurationSection smtpSettings)
        {
            foreach (DataRow row in dataSet.Tables[0].Rows)
            {

                string excelTitle = row["Excel_Title"].ToString();



 
                DataSet excelData = null;//這是EXCEL的DATASET



                //若之後要新增excel可在此插入判別式
                switch (excelTitle)
                {
                    case "":
                        excelData = ExcelHelper.();
                        break;

                    case "":
                        excelData = ExcelHelper.();
                        break;

                    case "":
                        excelData = ExcelHelper.();
                        break;

                    case "":
                        excelData = ExcelHelper.();
                        break;
                    // 添加更多的情况，如果有的话
                    case "":
                        excelData = ExcelHelper.();
                        break;
                    case "":
                        excelData = ExcelHelper.();
                        break;

                    default:
                        // 如果 excelTitle 不匹配任何情况的默认处理
                        // 可以选择抛出异常或执行其他逻辑
                        break;
                }



                DataSet emailData = EmailQuest(excelTitle);//這是EMAIL的DATASET


                if (excelData != null)
                {
                    List<string> actualFileNames = ExcelHelper.ExportDataSetToExcel(excelData, folderPath, excelTitle);

                    if (actualFileNames == null || actualFileNames.Count == 0)
                    {
                        LogError($"{dataSetName} Excel檔案生成失敗。");
                        Program.LogToFile($"{dataSetName} Excel檔案生成失敗。");
                        return;
                    }

                    foreach (var actualFileName in actualFileNames)
                    {
                        LogInformation($"{dataSetName} Excel檔案 {actualFileName} 成功生成。");

                        using (MailMessage message = new MailMessage())
                        using (FileStream fileStream = new FileStream(Path.Combine(folderPath, actualFileName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (SmtpClient client = ConfigureSmtpClient(smtpSettings))
                        {
                            ConfigureEmailMessage(message, smtpSettings, emailData, fileStream, actualFileName, excelTitle);

                            try
                            {
                                client.Send(message);
                                LogInformation($"{dataSetName} 郵件發送成功！");
                            }
                            catch (Exception ex)
                            {
                                LogError($"{dataSetName} 發送郵件時發生錯誤：{ex.Message}");
                            }
                        }
                    }
                }
                else
                {
                    LogError($"{dataSetName} 無效的 Excel 標題：{excelTitle}");
                }
            }
        }

        /// <summary>
        /// 郵件的主要設定
        /// </summary>
        /// <param name="message"></param>
        /// <param name="smtpSettings"></param>
        /// <param name="emailDataSet"></param>
        /// <param name="fileStream"></param>
        /// <param name="actualFileName"></param>
        /// <param name="excelTitle"></param>
        /// <exception cref="InvalidOperationException"></exception>
        private static void ConfigureEmailMessage(MailMessage message, IConfigurationSection smtpSettings, DataSet emailDataSet, FileStream fileStream, string actualFileName, string excelTitle)
        {
            // 1. 從配置中獲取發送郵件的地址
            string fromEmail = smtpSettings["Username"] ?? throw new InvalidOperationException("FromEmail 不可為空值");

            // 2. 獲取收件人列表
            List<string> toEmailList = new List<string>();
            foreach (DataRow row in emailDataSet.Tables[0].Rows)
            {
 
                string sendPerson = row["Send_Person"].ToString();

                toEmailList.Add(sendPerson);

            }
            string toEmails = string.Join(",", toEmailList);

            LogInformation($"生成的郵件地址： {toEmails}");

            // 3. 如果收件人地址為空，記錄錯誤並返回
            if (string.IsNullOrEmpty(toEmails))
            {
                LogError("ToEmail 不可為空值");
                return;
            }

            // 4. 將收件人地址添加到郵件的 "To" 集合中
            string[] recipients = toEmails.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var recipient in recipients)
            {
                try
                {
                    new MailAddress(recipient.Trim());
                    message.To.Add(new MailAddress(recipient.Trim()));
                }
                catch (FormatException)
                {
                    LogError($"無效的 Send_Person 或 CC_Person 郵件地址: {recipient}");
                }
            }

            // 5. 添加郵件附件
            Attachment emailAttachment = new Attachment(fileStream, actualFileName);
            message.Attachments.Add(emailAttachment);

            LogInformation("傳送檔案完成");

            // 6. 配置郵件主題、正文和其他信息
            foreach (DataRow row in emailDataSet.Tables[0].Rows)
            {
                string htmlContent = row["Email_html"].ToString();
                string emailTitle = row["Email_Title"].ToString();
                message.From = new MailAddress(fromEmail);
                // 7. 使用 excelTitle 生成主題
                message.Subject = string.IsNullOrEmpty(emailTitle) ? EmailSubject : emailTitle;

                // 8. 使用 EmailHtmlFilePath 或者從數據集中讀取正文內容
                message.Body = string.IsNullOrEmpty(htmlContent) ? File.ReadAllText(EmailHtmlFilePath) : htmlContent;
                message.IsBodyHtml = true;

                // 9. 添加抄送人
                string ccPersons = row["CC_Person"].ToString();
                if (!string.IsNullOrEmpty(ccPersons))
                {
                    string[] ccAddresses = ccPersons.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string ccPerson in ccAddresses)
                    {
                        try
                        {
                            new MailAddress(ccPerson.Trim());
                            message.CC.Add(new MailAddress(ccPerson.Trim()));
                        }
                        catch (FormatException)
                        {
                            LogError($"無效的 CC_Person 郵件地址: {ccPerson}");
                        }
                    }
                }
            }
        }


        public static DataSet EmailQuest()
        {
            string strsql = @"
sql
            ";


            return (DataSet)SQLDataAccess.ExecuteDataSet(SQLDataAccess.sql, CommandType.Text, strsql);

        }

        private static DataSet EmailQuest(string excelTitle)
        {
            string strsql = @"
        select (sql)
    ";

            // 使用 SqlParameter 防止 SQL 注入
            SqlParameter parameter = new SqlParameter("@TXTEXCEL", SqlDbType.NVarChar);
            parameter.Value = excelTitle;

            return (DataSet)SQLDataAccess.ExecuteDataSet(SQLDataAccess.sql, CommandType.Text, strsql, parameter);
        }

        private static SmtpClient ConfigureSmtpClient(IConfigurationSection smtpSettings)
        {
            return new SmtpClient(smtpSettings["SmtpServer"])
            {
                Port = int.TryParse(smtpSettings["Port"], out int port) ? port : 587,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(smtpSettings["Username"], smtpSettings["Password"]),
                EnableSsl = true
            };
        }

        private static void LogInformation(string message)
        {
            Console.WriteLine($"Information: {message}");
            LogToFile($"Information: {message}");
        }

        private static void LogError(string message)
        {
            Console.WriteLine($"Error: {message}");
            LogToFile($"Error: {message}");
        }
        public static void LogToFile(string log)
        {
            try
            {
                if (!Directory.Exists(LogFolderPath))
                {
                    Directory.CreateDirectory(LogFolderPath);
                }

                // 将日志追加到文件
                using (StreamWriter sw = File.AppendText(LogFilePath))
                {
                    sw.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {log}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"寫入LOG日誌時發生錯誤：{ex.Message}");
            }
        }

    }
}

