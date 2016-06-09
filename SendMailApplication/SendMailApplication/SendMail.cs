using Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SendMailApplication
{
    class SendMail
    {
        private const string LIST_OF_MAILS_FROM_FILE = @"D:\Mails.xlsx";
        private const string MSG_SUBJECT_ON_SUCC = "PNR Phase 3 Part 2 Internal Build Completed";
        private const string MSG_SUBJECT_ON_ERR = "Error in Build - PNR Phase 3 Part 2";
        private const string MSG_BODY_ON_SUCC = "Hi Team, \n\n          Latest code build is available for testing in the following url. Also excecuted BAT testcases, Please find the attachments for more details. \n\nURL: http://10.87.248.94:3232/Account/Login \nUserName: adm@releasephase32 \nPassword: 12345678  \n\nRegards,\nVinay";
        private const string MSG_BODY_ON_ERR = "Hi Team, \n\n          We are getting error while doing internal build. Please find the attached log files for more details. Respective owners please fix the build issue asap and let us know. \n\nRegards,\nVinay";

        static void Main(string[] args)
        {
            //PNR_BUILD_LOG file from jenkins workspace
            string[] lines = System.IO.File.ReadAllLines(@"D:\Deployment\LogFiles\log");
            string msgSubject = string.Empty;
            string msgBody = string.Empty;
            bool isRequireFix = false;
            List<string> ListOfMailIDs = new List<string>();

            //List of mail id to send mail - read from excel
            foreach (var worksheet in Workbook.Worksheets(LIST_OF_MAILS_FROM_FILE))
            {
                foreach (var row in worksheet.Rows)
                {
                    foreach (var cell in row.Cells)
                        ListOfMailIDs.Add(cell.Text);
                }
            }

            if (lines[lines.Length-1].Contains("SUCCESS"))
            {
                msgSubject = MSG_SUBJECT_ON_SUCC;
                msgBody = MSG_BODY_ON_SUCC;
            }
            else
            {
                msgSubject = MSG_SUBJECT_ON_ERR;
                msgBody = MSG_BODY_ON_ERR;
                isRequireFix = true;
            }

            //Method to get all log files in one place
            GetAllLogFiles();

            //Method to send mail - from outlook synced account
            SendMailsToUsers(ListOfMailIDs, msgSubject, msgBody, isRequireFix);

        }

        public static void GetAllLogFiles()
        {
            string log_source_path = string.Empty;
            string log_target_path = string.Empty;
            string map_sourceFile = string.Empty;
            string map_destFile = string.Empty;
            string log_file_name = string.Empty;

            //setting target path to combine all log files in one place
            log_target_path = @"D:\Deployment\App_For_SendingMail\LogFiles";

            //Copy PNR_LOG file from PNR Jenkins Workspace
            log_source_path = @"D:\Deployment\LogFiles";            
            map_sourceFile = System.IO.Path.Combine(log_source_path, "log");
            map_destFile = System.IO.Path.Combine(log_target_path, "PNR_APP_BUILD_LOG");
            System.IO.File.Copy(map_sourceFile, map_destFile, true);
            log_source_path = null; map_sourceFile = null; map_destFile = null;

            //Copy BAT_TEST_LOG file from Jenkins Workspace
            log_source_path = @"D:\Deployment\BuildTest_RobotFramwork\RobotUpdated\robo-dist\robo-dist\reports";
            map_sourceFile = System.IO.Path.Combine(log_source_path, "logs.html");
            map_destFile = System.IO.Path.Combine(log_target_path, "PNR_BAT_LOG.html");
            System.IO.File.Copy(map_sourceFile, map_destFile, true);
            log_source_path = null; map_sourceFile = null; map_destFile = null;


            //Copy DB_STRUCTURE_SCRIPT_EXEC_LOG file workspace
            log_source_path = @"D:\Deployment\Database_Files\Structure";
            map_sourceFile = System.IO.Path.Combine(log_source_path, "LOG.txt");
            map_destFile = System.IO.Path.Combine(log_target_path, "PNR_DB_STRUCTURE_LOG.txt");
            System.IO.File.Copy(map_sourceFile, map_destFile, true);
            log_source_path = null; map_sourceFile = null; map_destFile = null;

            //Copy DB_STOREDPROCEDURE_SCRIPT_EXEC_LOG file workspace
            log_source_path = @"D:\Deployment\Database_Files\StoredProcedures";
            map_sourceFile = System.IO.Path.Combine(log_source_path, "LOG.txt");
            map_destFile = System.IO.Path.Combine(log_target_path, "PNR_DB_STOREPROC_LOG.txt");
            System.IO.File.Copy(map_sourceFile, map_destFile, true);
            log_source_path = null; map_sourceFile = null; map_destFile = null;

            //Copy DB_TENANT_SCRIPT_EXEC_LOG file workspace
            log_source_path = @"D:\Deployment\Database_Files\Tenant";
            map_sourceFile = System.IO.Path.Combine(log_source_path, "LOG.txt");
            map_destFile = System.IO.Path.Combine(log_target_path, "PNR_DB_TENANT_DATASCRIPT_LOG.txt");
            System.IO.File.Copy(map_sourceFile, map_destFile, true);
            log_source_path = null; map_sourceFile = null; map_destFile = null;
        }

        public static void SendMailsToUsers(List<string> ListOfMailIDs, string msgSubject, string msgBody, bool isRequireFix)
        {
            Outlook._Application _app = null;
            Outlook.MailItem mail = null;
            try
            {
                _app = new Outlook.Application();
                mail = (Outlook.MailItem)_app.CreateItem(Outlook.OlItemType.olMailItem);
                string ids = string.Empty;
                foreach (string ip in ListOfMailIDs)
                   ids += ip + ";";
                
               // string ccMailIds = "saravanan_thoppae@infosys.com;Sathiya_Kailash@infosys.com;LNF_PNR_Build_Team@infosys.com";
                //string ccMailIds = "LNF_PNR_Build_Team@infosys.com";
                //mail.To = "MCTLNFRPNZ1Devs@infosys.com";
               // mail.CC = ccMailIds;

                mail.To = ids;
                // mail.To = ccMailIds;
                mail.Subject = msgSubject;
                mail.Body = msgBody;

                string[] filesAvailable = Directory.GetFiles(@"D:\Deployment\App_For_SendingMail\LogFiles");

                foreach (string fileName in filesAvailable)
                    mail.Attachments.Add(fileName);

                if (!isRequireFix)
                    mail.Importance = Outlook.OlImportance.olImportanceNormal; 
                else
                    mail.Importance = Outlook.OlImportance.olImportanceHigh;

                ((Outlook._MailItem)mail).Send();                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                _app = null;
                mail = null;
            }
        }

    }
}
