using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Windows.Forms;

namespace AttachmentSaver
{
    public partial class _Ribbon
    {
        Properties.Settings settings = Properties.Settings.Default;

        private void _Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            //On first initialization:
            if (settings.lastRun == DateTime.MinValue)
            {
                settings.lastRun = DateTime.Now;
                settings.Save();
            }
            else
            {
                UpdateInfo();
            }

            if (settings.autoEnabled)
            {
                autoCheck.Checked = true;
                autoCheck.Image = Properties.Resources.switchOn;
                runBtn.Enabled = false;
            }
        }

        private void AutoCheck_Click(object sender, RibbonControlEventArgs e)
        {
            if (autoCheck.Checked)
            {
                if (settings.opslagprofielen != "[]")
                {
                    autoCheck.Image = Properties.Resources.switchOn;
                    runBtn.Enabled = false;
                    settings.autoEnabled = true;
                    Globals.ThisAddIn.AddEventHandler_OnNewMailReceived();
                }
                else
                {
                    autoCheck.Checked = false;
                    MessageBox.Show("Stel eerst een opslagprofiel in om bijlagen op te slaan");
                }
            }
            else
            {
                autoCheck.Image = Properties.Resources.switchOff;
                runBtn.Enabled = true;
                settings.autoEnabled = false;
                Globals.ThisAddIn.RemoveEventHandler_OnNewMailReceived();
            }
            settings.Save();
        }

        private void ManageProfiles_Click(object sender, RibbonControlEventArgs e)
        {
            new SettingsWindow().Show();
            manageProfilesBtn.Enabled = false;
        }

        private void RunBtn_Click(object sender, RibbonControlEventArgs e)
        {
            if (settings.opslagprofielen != "[]")
            {
                runBtn.Enabled = false;
                Globals.ThisAddIn.SaveAttachments();
                runBtn.Enabled = true;
            }
            else
            {
                MessageBox.Show("Stel eerst een opslagprofiel in om bijlagen op te slaan");
            }
        }

        public void UpdateInfo()
        {
            string text = $"Laatst geïmporteerd om: {settings.lastRun}\n{settings.lastSuccessful} bestanden opgeslagen. ";
            if (settings.lastFailed > 0)
                text += settings.lastFailed + "bestanden gefaald.";

            setLastRunBtn.Label = text;
        }

        private void ViewLogBtn_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start(Directory.GetCurrentDirectory() + "\\log.txt");
        }

        private void setLastRunBtn_Click(object sender, RibbonControlEventArgs e)
        {
            new RunDatepicker().Show();
        }
    }
}
