using System;
using System.IO;
using System.Windows.Forms;

namespace EML_to_PST
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // Prevents resizing
            this.MaximizeBox = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Disable the button to prevent multiple clicks during conversion
            button1.Enabled = false;

            string sourceFolder = txtFolderPath.Text;
            string pstPath = txtPSTPath.Text;

            if (string.IsNullOrEmpty(sourceFolder) || string.IsNullOrEmpty(pstPath))
            {
                LogMessage("Molimo izaberite i izvorni folder i odredišnu datoteku.");
                return;
            }

            // Check if source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                LogMessage("Navedeni izvorni folder ne postoji.");
                return;
            }

            // Check if the source folder contains any .eml files
            var emlFiles = Directory.GetFiles(sourceFolder, "*.eml");
            if (emlFiles.Length == 0)
            {
                LogMessage("U navedenom izvornom folderu nisu pronađene EML datoteke.");
                return;
            }

            // Check if the output file path is valid
            if (!Directory.Exists(Path.GetDirectoryName(pstPath)))
            {
                LogMessage("Navedena putanja za PST datoteku je nevažeća.");
                return;
            }

            // Check if the PST file already exists
            if (File.Exists(pstPath))
            {
                // Check if the PST file already exists
                if (File.Exists(pstPath))
                {
                    var result = MessageBox.Show(
                        "Navedena PST datoteka već postoji. Da li želite da je zamenite?",
                        "Datoteka već postoji",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);

                    if (result == DialogResult.No)
                    {
                        LogMessage("Operacija je otkazana od strane korisnika.");
                        button1.Enabled = true;
                        return;
                    }

                    try
                    {
                        // Attempt to delete the existing file if user agrees to overwrite
                        File.Delete(pstPath);
                    }
                    catch (Exception ex)
                    {
                        LogMessage("Neuspešno brisanje postojeće PST datoteke: " + ex.Message);
                        return;
                    }
                }
            }

            progressBar1.Maximum = emlFiles.Length;
            progressBar1.Value = 0;
            label3.Text = $"0/{emlFiles.Length}";
            LogMessage("Početak konverzije...");

            // Initialize the PST file
            var personalStorage = Aspose.Email.Storage.Pst.PersonalStorage.Create(pstPath, Aspose.Email.Storage.Pst.FileFormatVersion.Unicode);
            var inboxFolder = personalStorage.RootFolder.AddSubFolder("Inbox");

            int convertedCount = 0;

            // Loop through each EML file, update the progress bar, and log messages
            foreach (var emlFile in emlFiles)
            {
                try
                {
                    var mailMessage = Aspose.Email.MailMessage.Load(emlFile);
                    inboxFolder.AddMessage(Aspose.Email.Mapi.MapiMessage.FromMailMessage(mailMessage));

                    // Log success for each file and update progress
                    LogMessage("Konvertovano: " + Path.GetFileName(emlFile));

                    convertedCount++;
                    progressBar1.Value = convertedCount;
                    label3.Text = $"{convertedCount}/{emlFiles.Length}";
                }
                catch (Exception ex)
                {
                    LogMessage("Greška pri konverziji " + Path.GetFileName(emlFile) + ": " + ex.Message);
                }
            }

            LogMessage("Konverzija je uspešno završena.");

            // Re-enable the button after conversion is complete
            button1.Enabled = true;
        }

        // Function to log messages to the text area
        private void LogMessage(string message)
        {
            // Append the message with a new line
            richTextBox1.AppendText(message + Environment.NewLine);
            // Set the selection to the end of the text
            richTextBox1.SelectionStart = richTextBox1.Text.Length;
            // Scroll to the caret to ensure the latest message is visible
            richTextBox1.ScrollToCaret();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFolderPath.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (var saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "PST Datoteke|*.pst";
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    txtPSTPath.Text = saveDialog.FileName;
                }
            }
        }

        private void infoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Create an instance of the Info form and show it
            InfoForm infoForm = new InfoForm();
            infoForm.ShowDialog(); // Show as a dialog box
        }
    }
}
