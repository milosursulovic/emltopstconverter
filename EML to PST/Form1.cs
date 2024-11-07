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
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sourceFolder = txtFolderPath.Text;
            string pstPath = txtPSTPath.Text;

            if (string.IsNullOrEmpty(sourceFolder) || string.IsNullOrEmpty(pstPath))
            {
                LogMessage("Please select both a source folder and a destination file.");
                return;
            }

            // Get all EML files in the selected folder
            var emlFiles = Directory.GetFiles(sourceFolder, "*.eml");
            progressBar1.Maximum = emlFiles.Length;
            progressBar1.Value = 0;
            LogMessage("Starting conversion...");

            // Initialize the PST file
            var personalStorage = Aspose.Email.Storage.Pst.PersonalStorage.Create(pstPath, Aspose.Email.Storage.Pst.FileFormatVersion.Unicode);
            var inboxFolder = personalStorage.RootFolder.AddSubFolder("Inbox");

            // Loop through each EML file, update the progress bar, and log messages
            foreach (var emlFile in emlFiles)
            {
                try
                {
                    var mailMessage = Aspose.Email.MailMessage.Load(emlFile);
                    inboxFolder.AddMessage(Aspose.Email.Mapi.MapiMessage.FromMailMessage(mailMessage));

                    // Log success for each file and update progress
                    LogMessage("Converted: " + Path.GetFileName(emlFile));
                    progressBar1.Value += 1;
                }
                catch (Exception ex)
                {
                    LogMessage("Error converting " + Path.GetFileName(emlFile) + ": " + ex.Message);
                }
            }

            LogMessage("Conversion completed successfully.");
        }

        // Function to log messages to the text area
        private void LogMessage(string message)
        {
            richTextBox1.AppendText(message + Environment.NewLine);
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
                saveDialog.Filter = "PST Files|*.pst";
                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    txtPSTPath.Text = saveDialog.FileName;
                }
            }
        }
    }
}
