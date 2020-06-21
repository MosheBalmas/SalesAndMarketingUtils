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
using SalesAndMarketingUtilsServices;

namespace SalesAndMarketingUtilsForm
{
    public partial class SaveAsForm : Form
    {
        byte[] bin;

        public SaveAsForm()
        {
            InitializeComponent();
            var srv = new SalesAndMarketingUtilsServices.SalesAndMarketingUtilsServices();
            //bin = srv.CreateSpreadsheet();
            textBoxFileName.Text =  string.Concat("JD_Plan", DateTime.Now.ToString("yyyyMMdd_HHmmss"), ".xlsx");

        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string spreadsheetPath = string.Concat("JD_Plan", DateTime.Now.ToString("yyyyMMdd_HHmmss"), ".xlsx");

            //create a SaveFileDialog instance with some properties
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Save Excel sheet";
            saveFileDialog1.Filter = "Excel files|*.xlsx|All files|*.*";
            saveFileDialog1.FileName = spreadsheetPath;

            //check if user clicked the save button
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //write the file to the disk
                File.WriteAllBytes(saveFileDialog1.FileName, bin);

            }

            Environment.Exit(0);
        }

        private void textBoxFileName_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
