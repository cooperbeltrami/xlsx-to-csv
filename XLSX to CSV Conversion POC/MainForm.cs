using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XLSX_to_CSV_Conversion_POC
{
    public partial class MainForm : Form
    {

        public MainForm()
        {

            InitializeComponent();
        }

        private void ConvertBtn_Click(object sender, EventArgs e)
        {
            // Create new open file dialog
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                RestoreDirectory = true,
                InitialDirectory = Environment.SpecialFolder.Desktop.ToString(),
                Title = "Browse for XLSX files to convert",
                DefaultExt = "xlsx",
                CheckFileExists = true,
                CheckPathExists = true,
                Multiselect = false, // This can be changed in the future
            };

            // Once user has chosen and click ok on the file dialog window
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Retreive info from select file
                string filePath = openFileDialog.FileName;

                using (var fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var fileReader = ExcelDataReader.ExcelReaderFactory.CreateReader(fileStream))
                    {
                        string csvText = "";

                        // Read each line of the file
                        do
                        {
                            int numberOfColumns = fileReader.FieldCount;

                            while (fileReader.Read()) // Each row
                            {
                                for (int column = 0; column < numberOfColumns; column++) // Each column 
                                {
                                    csvText += fileReader.GetValue(column) + ",";
                                }

                                csvText += "\n";
                            }
                        } while (fileReader.NextResult()); // For each sheet

                        // Print in textbox
                        textBox1.Text = csvText;

                       string newPath = filePath.Substring(0, filePath.IndexOf(".")) + ".csv";
                       File.AppendAllText(newPath, csvText);
                    }
                }
            }
        }
    }
}
