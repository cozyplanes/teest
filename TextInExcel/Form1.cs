using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TextInExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e) { }

        private void button1_Click(object sender, EventArgs e)
        {
            var filePath = Environment.ExpandEnvironmentVariables("%appdata%") + "\\Microsoft\\Excel\\XLSTART\\message.txt";
            //this gives C:\Users\<username>\AppData\Roaming\Microsoft\Excel\XLSTART\message.txt

            TextWriter txt = new StreamWriter(filePath);
            txt.Write(textBox1.Text);
            txt.Close();
            
            DialogResult confirm = MessageBox.Show(
                "Your message has been saved!\nClick OK to close this window, click Cancel to open saved path in explorer.",
                "Confirmation",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.None
            );
            
            if (confirm == DialogResult.Cancel)
            {
                // combine the arguments together
                // it doesn't matter if there is a space after ','
                string argument = "/select, \"" + filePath + "\"";

                Process.Start("explorer.exe", argument);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult about = MessageBox.Show(
                "Copyright (C) 2018 cozyplanes\nAll rights reserved\n\nThis program is protected under MIT License\n\nClick OK to close this window, click Cancel to view the source code in GitHub\n",
                "About",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.None
            );
            
            if (about == DialogResult.Cancel)
            {
                Process.Start("https://www.github.com/cozyplanes/teest");
            }
        }
    }
}
 