using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;

namespace AttachmentSaver
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace ons;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        Properties.Settings settings = Properties.Settings.Default;
        List<OpslagProfiel> profiles;

        Outlook.ItemsEvents_ItemAddEventHandler onNewMailReceived;

        List<string> log = new List<string>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ons = Application.GetNamespace("MAPI");
            inbox = ons.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            items = inbox.Items;

            onNewMailReceived = new Outlook.ItemsEvents_ItemAddEventHandler(OnNewMailReceived);
            if (settings.autoEnabled)
                AddEventHandler_OnNewMailReceived();
        }

        public void AddEventHandler_OnNewMailReceived()
        {
            items.ItemAdd += onNewMailReceived;
            SaveAttachments();
        }

        public void RemoveEventHandler_OnNewMailReceived()
        {
            items.ItemAdd -= onNewMailReceived;
        }

        void OnNewMailReceived(object item)
        {
            profiles = JsonConvert.DeserializeObject<List<OpslagProfiel>>(settings.opslagprofielen);

            try
            {
                var mail = (Outlook.MailItem)item;

                var resultSet = SaveAttachmentsOfMail(mail);

                settings.lastSuccessful = resultSet["successful"];
                settings.lastFailed += resultSet["failed"];
                settings.lastRun = DateTime.Now;
                settings.Save();
                Globals.Ribbons._Ribbon.UpdateInfo();
                SaveLog();
            }
            catch (InvalidCastException)
            {
            }
        }

        public void SaveAttachments()
        {
            profiles = JsonConvert.DeserializeObject<List<OpslagProfiel>>(settings.opslagprofielen);

            int successfullySavedFiles = 0;
            int failedFiles = 0;

            foreach (var item in items)
            {
                try
                {
                    var mail = (Outlook.MailItem)item;
                    if (mail.ReceivedTime > settings.lastRun)
                    {
                        var resultSet = SaveAttachmentsOfMail(mail);
                        successfullySavedFiles += resultSet["successful"];
                        failedFiles += resultSet["failed"];
                    }
                }
                catch (InvalidCastException)
                {
                    continue;
                }
            }

            settings.lastRun = DateTime.Now;
            settings.lastSuccessful = successfullySavedFiles;
            settings.lastFailed = failedFiles;
            settings.Save();
            Globals.Ribbons._Ribbon.UpdateInfo();
            SaveLog();
        }

        private Dictionary<string, int> SaveAttachmentsOfMail(Outlook.MailItem mail)
        {
            var resultSet = new Dictionary<string, int>();

            resultSet.Add("successful", 0);
            resultSet.Add("failed", 0);

            if (mail.Attachments.Count > 0)
            {
                foreach (var profile in profiles)
                {
                    if (profile.Sender != mail.SenderEmailAddress)
                        continue;

                    foreach (var attachment in mail.Attachments.Cast<Outlook.Attachment>())
                    {
                        // Check if attachment is not embedded (picture, signature, or other html element)
                        if (mail.HTMLBody.Contains(attachment.DisplayName) || string.IsNullOrEmpty(Path.GetExtension(attachment.FileName)))
                            continue;

                        string fileExtension = Path.GetExtension(attachment.FileName);
                        string filename = attachment.FileName.Substring(0, attachment.FileName.LastIndexOf(fileExtension));

                        if (!string.IsNullOrEmpty(profile.FilenameFilter) && !Regex.IsMatch(filename, profile.FilenameFilter))
                            continue;

                        if (profile.FileExtensions.Any() && !profile.FileExtensions.Contains(fileExtension))
                            continue;

                        string path = profile.FileLocation + "\\" + attachment.FileName;
                        try
                        {
                            int i = 1;
                            while (File.Exists(path))
                            {
                                if (Regex.Match(path, @"\s\(\d\)\.\w+$").Success)
                                    path = path.Remove(path.LastIndexOf(fileExtension) - 4, 4);
                                path = path.Insert(path.LastIndexOf(fileExtension), $" ({i})");
                                i++;
                            }

                            attachment.SaveAsFile(path);
                            log.Add("Succesfully saved file: " + path);
                            resultSet["successful"]++;
                        }
                        catch (Exception ex)
                        {
                            log.Add($"Failed to save file to: {path}\n\t{ex.GetType()}: {ex.Message}");
                            resultSet["failed"]++;
                        }
                    }
                }
            }

            return resultSet;
        }

        public void SaveLog()
        {
            try
            {
                using (StreamWriter w = File.AppendText("log.txt"))
                {
                    foreach (string message in log)
                        w.WriteLine(DateTime.Now + " - " + message);

                    log.Clear();
                }
            }
            catch (Exception)
            {
            }
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
        }
        
        #endregion
    }
}
