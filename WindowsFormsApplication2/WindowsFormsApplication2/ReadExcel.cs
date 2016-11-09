using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
<<<<<<< HEAD
using System.Text.RegularExpressions;
=======
using System.Linq;
using System.Text;
using System.Threading.Tasks;
>>>>>>> parent of f9987bf... updated
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml;

namespace WindowsFormsApplication2
{
    public partial class ReadExcel : Form
    {

        string filePath = string.Empty;
        string fileExt = string.Empty;
        string con;
        OleDbConnection conn;
        OleDbCommand oconn;
        OleDbDataAdapter sda;
<<<<<<< HEAD
        int numEntries;
        string filePath = string.Empty;
        String worksheetName;
        Decimal result;

=======
        DataTable data;
>>>>>>> parent of f9987bf... updated

        public ReadExcel()
        {
            InitializeComponent();
        }

        private void ChooseFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
<<<<<<< HEAD
                excelFilePath = file.FileName;
                fileExt = Path.GetExtension(excelFilePath);
                Cursor = Cursors.WaitCursor;
=======
                filePath = file.FileName;
                fileExt = Path.GetExtension(filePath);
>>>>>>> parent of f9987bf... updated
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        if (fileExt.CompareTo(".xls") == 0)
                            con = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                        else
                            con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';"; //for above excel 2007  

                        Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Open(excelFilePath);
                        Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
                        worksheetName = worksheet.Name;
                        workbook.Close(0);
                        app.Quit();
                        String command = "SELECT * FROM [" + worksheetName + "$]";
                        conn = new OleDbConnection(con);
                        oconn = new OleDbCommand(command, conn);
                        conn.Open();
                        sda = new OleDbDataAdapter(oconn);
                        DataTable data = new DataTable();
                        sda.Fill(data);
                        DataRow row = data.Rows[0];
                        data.Rows.Remove(row);
                        dataGridView2.DataSource = data;
                        Cursor = Cursors.Default;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                }
            }
        }

        private void CloseBtn_Click(object sender, EventArgs e)
        {
            this.Close(); //to close the window(Form1)
        }

        private void ConvXMLbtn_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count == 0)
            {
                MessageBox.Show("Please load an Excel file before attempting to convert to XML.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                saveFileDialog1.DefaultExt = "xml";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    saveFileDialog1_FileOk();
                }
            }
        }

        private void saveFileDialog1_FileOk()
        {
            string name = saveFileDialog1.FileName;

            string[,] tempDataTable = new string[dataGridView2.Rows.Count, 8];
            int rowNumber = 0;
            

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                tempDataTable[rowNumber, 0] = row.Cells[1].Value.ToString();
                tempDataTable[rowNumber, 1] = row.Cells[7].Value.ToString();
                tempDataTable[rowNumber, 2] = row.Cells[8].Value.ToString();
                tempDataTable[rowNumber, 3] = row.Cells[9].Value.ToString();
                tempDataTable[rowNumber, 4] = row.Cells[10].Value.ToString();
                tempDataTable[rowNumber, 5] = row.Cells[11].Value.ToString();
                tempDataTable[rowNumber, 6] = row.Cells[12].Value.ToString();
                tempDataTable[rowNumber, 7] = row.Cells[13].Value.ToString();
                rowNumber++;
            }
<<<<<<< HEAD

            float numFiles = (tempDataTable.GetLength(0) / numEntries);
            int counter; 
            int rows = tempDataTable.GetLength(0);

            for (int x = 0; x < numFiles+1; x++)
=======
            using (XmlWriter writer = XmlWriter.Create(name))
>>>>>>> parent of f9987bf... updated
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("SetupInvMaster");
                for (int i = 0; i < tempDataTable.GetLength(0); i++)
                {
<<<<<<< HEAD
                    writer.WriteStartDocument();
                    writer.WriteStartElement("SetupInvMaster");
                    for (int i = 0; i < numEntries; i++)
                    {
                        counter = i + x * numEntries;

                        if (tempDataTable[counter, 1].Equals("StockUom"))
                        {

                        }
                        else
                        {
                            writer.WriteStartElement("Item");
                            writer.WriteStartElement("Key");
                            writer.WriteElementString("StockCode", tempDataTable[counter, 0]);
                            writer.WriteEndElement();
                            writer.WriteElementString("StockUom", tempDataTable[counter, 1]);
                            writer.WriteElementString("AlternateUom", tempDataTable[counter, 2]);
                            writer.WriteElementString("OtherUom", tempDataTable[counter, 3]);
                            writer.WriteElementString("ConvFactAltUom", tempDataTable[counter, 4]);
                            writer.WriteElementString("ConvMulDiv", tempDataTable[counter, 5]);
                            writer.WriteElementString("ConvFactOthUom", tempDataTable[counter, 6]);
                            writer.WriteElementString("MulDiv", tempDataTable[counter, 7]);
                            writer.WriteEndElement();
                        }

                        if (counter+2 > rows)
                        {
                            i = numEntries + 10;
                        }

                    }
=======
                    writer.WriteStartElement("Item");
                    writer.WriteStartElement("Key");
                    writer.WriteElementString("StockCode", tempDataTable[i,0]);
                    writer.WriteEndElement();
                    writer.WriteElementString("StockUom", tempDataTable[i, 1]);
                    writer.WriteElementString("AlternateUom", tempDataTable[i, 2]);
                    writer.WriteElementString("OtherUom", tempDataTable[i, 3]);
                    writer.WriteElementString("ConvFactAltUom", tempDataTable[i, 4]);
                    writer.WriteElementString("ConvMulDiv", tempDataTable[i, 5]);
                    writer.WriteElementString("ConvFactOthUom", tempDataTable[i, 6]);
                    writer.WriteElementString("MulDiv", tempDataTable[i, 7]);
>>>>>>> parent of f9987bf... updated
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
                writer.WriteEndDocument();
                MessageBox.Show("Generation of XML file has finished.","Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ConvXMLbtn2_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count == 0)
            {
                MessageBox.Show("Please load an Excel file before attempting to convert to XML.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                saveFileDialog1.DefaultExt = "xml";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    saveFileDialog2_FileOk();
                }
            }
        }

        private void saveFileDialog2_FileOk()
        {
            Cursor = Cursors.WaitCursor;
            string name = saveFileDialog1.FileName;
<<<<<<< HEAD
            string nameNoExt = Path.GetFileNameWithoutExtension(name);
            filePath = Path.GetDirectoryName(name);
            //conn = new OleDbConnection(con);
            //Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Open(excelFilePath);
            //Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets[1] as Microsoft.Office.Interop.Excel.Worksheet;
            //String columnName = worksheet.Columns[110].Address;
            //columnName = columnName.Substring(1,2);
            //String command = "SELECT * FROM [" + worksheetName + "$]";
            //conn = new OleDbConnection(con);
            //oconn = new OleDbCommand(command, conn);
            //conn.Open();
            //sda = new OleDbDataAdapter(oconn);
            //workbook.Close(0);
            //app.Quit();
            //DataTable data = new DataTable();
            //sda.Fill(data);
            //dataGridView2.DataSource = data;
=======
>>>>>>> parent of f9987bf... updated

            string[,] tempDataTable = new string[dataGridView2.Rows.Count, 3];
            int rowNumber = 0;
            string checker;

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
<<<<<<< HEAD
                checker = row.Cells[109].Value.ToString();
                //if (!string.IsNullOrEmpty(checker))
                decimal.TryParse(checker, out result);
                MessageBox.Show(checker);

                if (!string.IsNullOrEmpty(checker))
                {
                        tempDataTable[rowNumber, 0] = row.Cells[3].Value.ToString();
                        tempDataTable[rowNumber, 1] = row.Cells[1].Value.ToString();
                        tempDataTable[rowNumber, 2] = row.Cells[109].Value.ToString();
                        rowNumber++;
                }
            }

            float numFiles = (rowNumber / numEntries);
            int counter;
            int rows = tempDataTable.GetLength(0);

            for (int x = 0; x < numFiles + 1; x++)
=======
                    tempDataTable[rowNumber, 0] = row.Cells["Warehouse"].Value.ToString();
                    tempDataTable[rowNumber, 1] = row.Cells["StockCode"].Value.ToString();
                    tempDataTable[rowNumber, 2] = row.Cells["Cost"].Value.ToString();
                    rowNumber++;
            }

            using (XmlWriter writer = XmlWriter.Create(name))
>>>>>>> parent of f9987bf... updated
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("PostInvCostChange");
                for (int i = 0; i < tempDataTable.GetLength(0); i++)
                {
                    writer.WriteStartElement("Item");
                    writer.WriteElementString("Warehouse", tempDataTable[i, 0]);
                    writer.WriteElementString("StockCode", tempDataTable[i, 1]);
                    writer.WriteElementString("NewUnitCost", tempDataTable[i, 2]);
                    writer.WriteElementString("UpdateLastCost", "Y");
                    writer.WriteElementString("Reference", "UOM CHG");
                    writer.WriteElementString("Notation", "Cost change Note");
                    writer.WriteEndElement();
                }
<<<<<<< HEAD
            }
            Cursor = Cursors.Default;
            MessageBox.Show("Generation of XML file has finished.", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (filePath != string.Empty)
            {
                Process.Start("explorer.exe", @filePath);
            }
            else
            {
                MessageBox.Show("Please convert an XML file before attempting to open the directory.", "Error.", MessageBoxButtons.OK, MessageBoxIcon.Error);
=======
                writer.WriteEndElement();
                writer.WriteEndDocument();
                MessageBox.Show("Generation of XML file has finished.", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);
>>>>>>> parent of f9987bf... updated
            }
        }
    }
}

