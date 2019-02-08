using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {

        string path = Path.Combine(Directory.GetCurrentDirectory(), "APM Members.cfg");

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load_1(object sender, EventArgs e)
        {
            var lines = File.ReadAllLines(Path.Combine(Directory.GetCurrentDirectory(), "APM Members.cfg"));
            for (var i = 0; i < lines.Length; i += 1)
            {
                var line = lines[i];
                Form1.ApmList.Add(line.ToString());
                rtApmList.Text = rtApmList.Text + line +"\n";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            File.AppendAllLines(path,Form1.ApmList);
            //File.Create(path);
            Form1.ApmList.Remove("\n");
            rtApmList.SaveFile(path, RichTextBoxStreamType.PlainText);
            MessageBox.Show("Saved!");
        }
    }
}
