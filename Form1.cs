using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace TemporalBind
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();                   
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if (File.Exists(openFileDialog1.FileName)) {
                textBox1.Text = openFileDialog1.FileName;                
                button3.Enabled = true;
                textBox5.Text = "0";
                label1.Text = "/1";
            }
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            openFileDialog2.ShowDialog();
            if (File.Exists(openFileDialog2.FileName))
            {
                textBox2.Text = openFileDialog2.FileName;                
                button2.Enabled = true;
                button4.Enabled = true;
                textBox5.Text = "0";
                label1.Text = "/2";
            }
            
        }
       private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog3.ShowDialog();
            if (File.Exists(openFileDialog3.FileName))
            {
                textBox3.Text = openFileDialog3.FileName;
                button5.Enabled = true;
                textBox5.Text = "0";
                label1.Text = "/3";
            }
        }
        private void button5_Click_1(object sender, EventArgs e)
        {
            openFileDialog4.ShowDialog();
            if (File.Exists(openFileDialog3.FileName))
            {
                textBox4.Text = openFileDialog4.FileName;                
                textBox5.Text = "0";
                label1.Text = "/4";
            }
            
        }

        void generator() {
            var excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.Workbooks.Add();
            var ws = (Microsoft.Office.Interop.Excel.Worksheet)excelApp.ActiveSheet;
            ws.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlNoSelection;
            var file1str = System.IO.File.ReadAllText(openFileDialog1.FileName);
            file1str = file1str.Replace("\u0000", "").Replace("\u0009", "");
            var row = 2;
            var timestamp = 1;
            var index = 2;
            var content = 3;
            ws.Columns[1].ColumnWidth = 23;
            ws.Columns[2].ColumnWidth = 8;
            /*
            ws.Cells[1, 1] = "Timestamp";
            ws.Cells[1, 2] = "Index";
            ws.Cells[1, content] = openFileDialog1.SafeFileName;
            */
            textBox5.Text = "1";
            progressBar1.Maximum = file1str.Split('\n').Length+1;
            progressBar1.Minimum = 0;
            foreach (var text in file1str.Split('\n'))
            {
                var splitText = text.Split(' ');
                if (splitText.Length > 0)
                {
                    ws.Cells[row, timestamp] = splitText[0]; // Make sure it is timestamp
                    ws.Cells[row, index] = string.Format(String.Format("{0:D10}", row));
                    if (splitText.Length > 1)
                    {
                        splitText[0] = "";
                        string combinedString = "";
                        foreach (var tx in splitText)
                        {
                            combinedString += " " + tx;
                        }
                        ws.Cells[row, content] = combinedString.ToString();
                    }
                }
                else
                {
                    break;
                }
                progressBar1.Value = row;
                row++;
            }
            ws.Columns[content].ColumnWidth = 80;
            content++;

            if (textBox2.Text != "" && textBox2.Text != textBox1.Text)
            {
                textBox5.Text = "2";
                file1str = System.IO.File.ReadAllText(textBox2.Text);
                file1str = file1str.Replace("\u0000", "").Replace("\u0009", "");
                //ws.Cells[1, content] = openFileDialog2.SafeFileName;
                progressBar1.Maximum = file1str.Split('\n').Length + row;
                progressBar1.Minimum = row;
                foreach (var text in file1str.Split('\n'))
                {
                    progressBar1.Value = row; ;
                    var splitText = text.Split(' ');
                    if (splitText.Length > 0)
                    {
                        ws.Cells[row, timestamp] = splitText[0]; // Make sure it is timestamp
                        ws.Cells[row, index] = string.Format(String.Format("{0:D10}", row));
                        if (splitText.Length > 1)
                        {
                            splitText[0] = "";
                            string combinedString = "";
                            foreach (var tx in splitText)
                            {
                                combinedString += " " + tx;
                            }
                            ws.Cells[row, content] = combinedString.ToString();
                        }
                    }
                    else
                    {
                        break;
                    }
                    row++;
                }
                ws.Columns[content].ColumnWidth = 80;
                content++;
            }

            if (File.Exists(textBox3.Text)) {
                if (textBox3.Text != "" && textBox2.Text != textBox3.Text && textBox1.Text != textBox3.Text)                    
                {
                    file1str = System.IO.File.ReadAllText(textBox3.Text);
                    file1str = file1str.Replace("\u0000", "").Replace("\u0009", "");
                    //ws.Cells[1, content] = openFileDialog3.SafeFileName;
                    progressBar1.Maximum = file1str.Split('\n').Length + row;
                    progressBar1.Minimum = row;
                    foreach (var text in file1str.Split('\n'))
                    {
                        progressBar1.Value = row;
                        var splitText = text.Split(' ');
                        if (splitText.Length > 0)
                        {
                            ws.Cells[row, timestamp] = splitText[0]; // Make sure it is timestamp
                            ws.Cells[row, index] = string.Format(String.Format("{0:D10}", row));
                            if (splitText.Length > 1)
                            {
                                splitText[0] = "";
                                string combinedString = "";
                                foreach (var tx in splitText)
                                {
                                    combinedString += " " + tx;
                                }
                                ws.Cells[row, content] = combinedString.ToString();
                            }
                        }
                        else
                        {
                            break;
                        }
                        row++;
                    }
                }
                ws.Columns[content].ColumnWidth = 80;
                content++;
            }

            if (File.Exists(textBox4.Text)) {
                if (textBox4.Text != "" && textBox3.Text != textBox4.Text && textBox2.Text != textBox4.Text && textBox1.Text != textBox4.Text)
                {
                    textBox5.Text = "4";
                    file1str = System.IO.File.ReadAllText(textBox4.Text);
                    file1str = file1str.Replace("\u0000", "").Replace("\u0009", "");
                    //ws.Cells[1, content] = openFileDialog4.SafeFileName;
                    progressBar1.Maximum = file1str.Split('\n').Length + row;
                    progressBar1.Minimum = row;
                    foreach (var text in file1str.Split('\n'))
                    {
                        progressBar1.Value = row;
                        var splitText = text.Split(' ');
                        if (splitText.Length > 0)
                        {
                            ws.Cells[row, timestamp] = splitText[0]; // Make sure it is timestamp
                            ws.Cells[row, index] = string.Format(String.Format("{0:D10}", row));
                            if (splitText.Length > 1)
                            {
                                splitText[0] = "";
                                string combinedString = "";
                                foreach (var tx in splitText)
                                {
                                    combinedString += " " + tx;
                                }
                                ws.Cells[row, content] = combinedString.ToString();
                            }
                        }
                        else
                        {
                            break;
                        }
                        row++;
                    }
                }
            }

            ws.Columns[content].ColumnWidth = 80;
            dynamic allData = ws.UsedRange;
            allData.Sort(allData.Columns[1], Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending);
            //ws.Sort(ws.Columns[1], Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending);

            var k = ws.Range[ws.Cells[1, 1], ws.Cells[row, content]];
            MessageBox.Show("Temporal Merge Completed.");
            progressBar1.Value = progressBar1.Minimum;            
            excelApp.Visible = true;
            button1.Enabled = true;
            button2.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            generator();
        }


    }
}
