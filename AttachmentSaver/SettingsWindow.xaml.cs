using Microsoft.Office.Tools.Ribbon;
using System.Windows;
using System.Windows.Controls;
using System.Linq;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Windows.Input;
using System.IO;

namespace AttachmentSaver
{
    /// <summary>
    /// Interaction logic for SettingsWindow.xaml
    /// </summary>
    public partial class SettingsWindow : Window
    {
        Properties.Settings settings = Properties.Settings.Default;

        List<OpslagProfiel> profiles;
        int selected = -1; // Currently selected profile in the profiles listbox (default = -1)

        public SettingsWindow()
        {
            InitializeComponent();

            profiles = JsonConvert.DeserializeObject<List<OpslagProfiel>>(settings.opslagprofielen);

            foreach (var profile in profiles)
                profilesListbox.Items.Add(profile.Name);

            profilesListbox.Items.Add("<Nieuw>");
        }

        private void OnProfileSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int i = profilesListbox.SelectedIndex;

            if (i == -1)
                return;
            if (i == profilesListbox.Items.Count - 1) // <Nieuw> profile selected
            {
                if (selected > -1)
                    Empty();
            }   
            else
            {
                selected = i;

                nameTextbox.Text = profiles[i].Name;
                senderTextbox.Text = profiles[i].Sender;
                filenameTextbox.Text = profiles[i].FilenameFilter;
                fileLocationTextbox.Text = profiles[i].FileLocation;
                fileExtensionsListbox.Items.Clear();

                foreach (string fileExtension in profiles[i].FileExtensions)
                    fileExtensionsListbox.Items.Add(fileExtension);
            }
        }

        private void Empty()
        {
            nameTextbox.Text = "";
            fileLocationTextbox.Text = "";
            senderTextbox.Text = "";
            filenameTextbox.Text = "";
            fileExtensionsListbox.Items.Clear();

            selected = -1;
        }

        private void nameTextbox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (selected > -1)
            {
                profiles[selected].Name = nameTextbox.Text;
                profilesListbox.Items[selected] = nameTextbox.Text;
            }
        }

        private void senderTextbox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (selected > -1)
                profiles[selected].Sender = senderTextbox.Text;
        }

        private void filenameTextbox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (selected > -1)
                profiles[selected].FilenameFilter = filenameTextbox.Text;
        }

        private void fileExtensionTextbox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return) // On Enter press
            {
                AddFileExtension(sender, e);
                e.Handled = true;
                addFileExtension.Focus();
            }
        }

        private void AddFileExtension(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(fileExtensionTextbox.Text))
            {
                MessageBox.Show("Lege input", "Error");
                return;
            }
            else
            {
                string fileExtension = (fileExtensionTextbox.Text.First() == '.' ? "" : ".") + fileExtensionTextbox.Text;

                fileExtensionTextbox.Text = "";
                fileExtensionsListbox.Items.Add(fileExtension);

                if (selected > -1)
                    profiles[selected].FileExtensions.Add(fileExtension);
            }
        }

        private void RemoveFileExtension(object sender, RoutedEventArgs e)
        {
            fileExtensionsListbox.Items.RemoveAt(
                fileExtensionsListbox.Items.IndexOf(fileExtensionsListbox.SelectedItem));

            if (selected > -1)
                profiles[selected].FileExtensions.Remove((string)fileExtensionsListbox.SelectedItem);
        }

        private void Browse(object sender, RoutedEventArgs e)
        {
            var fileLocationDialog = new System.Windows.Forms.FolderBrowserDialog();

            if (fileLocationDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string fileLocation = fileLocationDialog.SelectedPath;
                fileLocationTextbox.Text = fileLocation;

                if (selected > -1)
                    profiles[selected].FileLocation = fileLocation;
            }
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            if (selected == -1 &&   // If all is empty, ignore.
                string.IsNullOrEmpty(nameTextbox.Text) &&
                string.IsNullOrEmpty(senderTextbox.Text) &&
                string.IsNullOrEmpty(filenameTextbox.Text) &&
                fileExtensionsListbox.Items.Count == 0 &&
                string.IsNullOrEmpty(fileLocationTextbox.Text))
                return;

            if (string.IsNullOrEmpty(nameTextbox.Text))
            {
                MessageBox.Show("Opslagprofiel moet een naam hebben", "Error");
                return;
            }

            if (string.IsNullOrEmpty(fileLocationTextbox.Text))
            {
                MessageBox.Show("Geen opslaglocatie gekozen", "Error");
                return;
            }

            if (!Directory.Exists(fileLocationTextbox.Text))
            {
                MessageBox.Show("Opslaglocatie bestaat niet", "Error");
                return;
            }

            var profile = new OpslagProfiel
            {
                Name = nameTextbox.Text,
                Sender = senderTextbox.Text,
                FilenameFilter = filenameTextbox.Text,
                FileExtensions = fileExtensionsListbox.Items.Cast<string>().ToList(),
                FileLocation = fileLocationTextbox.Text
            };

            if (selected == -1)
            {
                profiles.Add(profile);
                profilesListbox.Items.Insert(profilesListbox.Items.Count - 1, nameTextbox.Text);

                Empty();
            }
            else
            {
                profiles[selected] = profile;
            }

            settings.opslagprofielen = JsonConvert.SerializeObject(profiles);
            settings.Save();
        }

        private void Delete(object sender, RoutedEventArgs e)
        {
            if (selected > -1)
            {
                profilesListbox.Items.RemoveAt(
                    profilesListbox.Items.IndexOf(profilesListbox.SelectedItem));

                profiles.RemoveAt(selected);
                settings.opslagprofielen = JsonConvert.SerializeObject(profiles);
                settings.Save();

                Empty();
            }
        }

        private void Cancel(object sender, RoutedEventArgs e)
        {
            Hide();
            Globals.Ribbons._Ribbon.manageProfilesBtn.Enabled = true;
        }

        private void OnClose(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Globals.Ribbons._Ribbon.manageProfilesBtn.Enabled = true;
        }
    }
}
