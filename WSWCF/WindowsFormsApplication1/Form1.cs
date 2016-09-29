using System;
using System.IO;
using System.Net;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using MiWCF;
using Excel = Microsoft.Office.Interop.Excel;
using WindowsApp.WSElectricEnergy;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            webBrowser1.Navigate("https://www.cia.gov/library/publications/the-world-factbook/geos/by.html");
            dataGridView1.DataSource = CreateDataTable();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DisplayData();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            DownloadFile();
        }

        //Extract Data from website
        public EEData[] extractData()
        {
            //Copy URL's source code and download it as a string
            WebClient address = new WebClient();
            string url = address.DownloadString("https://www.cia.gov/library/publications/the-world-factbook/geos/by.html");

            //Create arrays of data to store and substrings to search on
            EEData[] data = new EEData[9];
            string[] eletricStrings = new String[9];

            //Get substring of Electricity related text, between the words Energy and oil
            eletricStrings[0] = getTextBetween(url, "Energy", "oil");

            //Use the electricity substring to find the first data
            data[0] = new EEData(getTextBetween(eletricStrings[0], ">Electricity - ", ":</a>"), getTextBetween(eletricStrings[0], "<div class=category_data>", "</div>"));

            //Search for new data in each new substring
            for (int i = 1; i < data.Length; i++)
            {
                eletricStrings[i] = eletricStrings[i - 1].Substring(eletricStrings[i - 1].IndexOf(data[i - 1].Data));
                data[i] = new EEData(getTextBetween(eletricStrings[i], ">Electricity - ", ":</a>"), getTextBetween(eletricStrings[i], "<div class=category_data>", "</div>"));
            }
            return data;
        }

        //Display Data on WebBrowser
        public void DisplayData()
        {
            EEData[] data = extractData();

            //Display data
            webBrowser1.DocumentText = "<html><body> " +
                "<h1>--Electricity--</h1> <hr> <br/>" +
                "<b>" + data[0].Type + ":</b> " + data[0].Data + "<br/><br/>" +
                "<b>" + data[1].Type + ":</b> " + data[1].Data + "<br/><br/>" +
                "<b>" + data[2].Type + ":</b> " + data[2].Data + "<br/><br/>" +
                "<b>" + data[3].Type + ":</b> " + data[3].Data + "<br/><br/>" +
                "<b>" + data[4].Type + ":</b> " + data[4].Data + "<br/><br/>" +
                "<b>" + data[5].Type + ":</b> " + data[5].Data + "<br/><br/>" +
                "<b>" + data[6].Type + ":</b> " + data[6].Data + "<br/><br/>" +
                "<b>" + data[7].Type + ":</b> " + data[7].Data + "<br/><br/>" +
                "<b>" + data[8].Type + ":</b> " + data[8].Data + "<br/><br/><hr>" +
                "</body></html>";
        }

        //Method to get a substring between a given start and end point from a source string
        public string getTextBetween(string sourceCode, string start, string end)
        {
            int begin;
            int finish;

            if (sourceCode.Contains(start) && sourceCode.Contains(end))
            {
                begin = sourceCode.IndexOf(start, 0) + start.Length;
                finish = sourceCode.IndexOf(end, begin);
                return sourceCode.Substring(begin, finish - begin);
            }
            else
            {
                return "";
            }
        }

        //Get Data inside a DataGridView
        public System.Data.DataTable CreateDataTable()
        {
            EEData[] data = extractData();

            System.Data.DataTable tableGrid = new System.Data.DataTable();
            tableGrid.Columns.Add("Type", typeof(string));
            tableGrid.Columns.Add("Data", typeof(string));
            for (int i = 0; i < data.Length; i++)
                tableGrid.Rows.Add(data[i].Type, data[i].Data);
            
            return tableGrid;
        }
        
        //Method to Ask User where to save the Excel File
        public void DownloadFile()
        {
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Excel Documents (*.xls)|*.xls";
            save.FileName = "Export-Electric-Energy-Data.xls";
            if (save.ShowDialog() == DialogResult.OK)
            {
                ExportToExcel(dataGridView1, save.FileName);
            }
        }

        //Method to create an Excel File from the DataGrid Table
        private void ExportToExcel(DataGridView tableGrid, string filename)
        {
            string stOutput = "";
            //export titles:
            string sHeaders = "";

            //get columns
            for (int j = 0; j < tableGrid.Columns.Count; j++)
                sHeaders = sHeaders.ToString() + Convert.ToString(tableGrid.Columns[j].HeaderText) + "\t";
            stOutput += sHeaders + "\r\n";
            //get rows
            for (int i = 0; i < tableGrid.RowCount - 1; i++)
            {
                string stLine = "";
                for (int j = 0; j < tableGrid.Rows[i].Cells.Count; j++)
                    stLine = stLine.ToString() + Convert.ToString(tableGrid.Rows[i].Cells[j].Value) + "\t";
                stOutput += stLine + "\r\n";
            }
            Encoding utf16 = Encoding.GetEncoding(1254);
            byte[] output = utf16.GetBytes(stOutput);
            FileStream fs = new FileStream(filename, FileMode.Create);
            BinaryWriter bw = new BinaryWriter(fs);
            bw.Write(output, 0, output.Length); 
            bw.Flush();
            bw.Close();
            fs.Close();
        }


    }
}
