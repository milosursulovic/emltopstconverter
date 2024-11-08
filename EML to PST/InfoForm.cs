using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EML_to_PST
{
    public partial class InfoForm : Form
    {
        public InfoForm()
        {
            InitializeComponent();

            // Form width
            int formWidth = this.ClientSize.Width;

            Label lblTitle = new Label();
            lblTitle.Text = "EML u PST Konverter";
            lblTitle.Font = new Font("Arial", 12, FontStyle.Bold);
            lblTitle.AutoSize = true;
            lblTitle.TextAlign = ContentAlignment.MiddleCenter; // Center text inside the label
            this.Controls.Add(lblTitle);

            // Recalculate label position after AutoSize has been set
            lblTitle.Location = new Point((formWidth - lblTitle.Width) / 2, 20);

            Label lblVersion = new Label();
            lblVersion.Text = "Verzija: 1.0.0";
            lblVersion.AutoSize = true;
            lblVersion.TextAlign = ContentAlignment.MiddleCenter; // Center text inside the label
            this.Controls.Add(lblVersion);
            lblVersion.Location = new Point((formWidth - lblVersion.Width) / 2, 60);

            Label lblDeveloper = new Label();
            lblDeveloper.Text = "Napravio: Miloš Ursulović";
            lblDeveloper.AutoSize = true;
            lblDeveloper.TextAlign = ContentAlignment.MiddleCenter; // Center text inside the label
            this.Controls.Add(lblDeveloper);
            lblDeveloper.Location = new Point((formWidth - lblDeveloper.Width) / 2, 100);

            Label lblDescription = new Label();
            lblDescription.Text = "Ovaj alat konvertuje EML fajlove u PST format radi lakšeg upravljanja emailovima u Outlook-u.";
            lblDescription.AutoSize = true;
            lblDescription.TextAlign = ContentAlignment.MiddleCenter; // Center text inside the label
            lblDescription.MaximumSize = new Size(250, 0); // Wrap text within 250 pixels
            this.Controls.Add(lblDescription);
            lblDescription.Location = new Point((formWidth - lblDescription.Width) / 2, 140);
        }
    }
}
