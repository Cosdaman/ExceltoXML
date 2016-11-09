using System;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace WindowsFormsApplication2
{
    public partial class ReadExcel : Form
    {

        string excelFilePath = string.Empty;
        string fileExt = string.Empty;
        string con;
        OleDbConnection conn;
        OleDbCommand oconn;
        OleDbDataAdapter sda;
        DataTable data;
        int numEntries;
        string filePath = string.Empty;
        String worksheetName;

        public ReadExcel()
        {
            InitializeComponent();
        }

        private void ChooseFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                excelFilePath = file.FileName;
                fileExt = Path.GetExtension(excelFilePath);
                Cursor = Cursors.WaitCursor;
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        if (fileExt.CompareTo(".xls") == 0)
                            con = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFilePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                        else
                            con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';"; //for above excel 2007  

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
            this.Close();
        }

        private void ConvXMLbtn_Click(object sender, EventArgs e)
        {
            
            Int32.TryParse(textBox1.Text, out numEntries);

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
            string nameNoExt = Path.GetFileNameWithoutExtension(name);
            filePath = Path.GetDirectoryName(name);
            string[,] tempDataTable = new string[dataGridView2.Rows.Count, 8];
            int rowNumber = 0;

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                //1st row of excel is the title of the columns; row.Cells[title]
                //ensure there are no duplicate column names in excel file
                tempDataTable[rowNumber, 0] = row.Cells["*StockCode"].Value.ToString();
                tempDataTable[rowNumber, 1] = row.Cells["StockUom"].Value.ToString();
                tempDataTable[rowNumber, 2] = row.Cells["AlternateUom"].Value.ToString();
                tempDataTable[rowNumber, 3] = row.Cells["OtherUom"].Value.ToString();
                tempDataTable[rowNumber, 4] = row.Cells["ConvFactAltUom"].Value.ToString();
                tempDataTable[rowNumber, 5] = row.Cells["ConvMulDiv"].Value.ToString();
                tempDataTable[rowNumber, 6] = row.Cells["ConvFactOthUom"].Value.ToString();
                tempDataTable[rowNumber, 7] = row.Cells["MulDiv"].Value.ToString();
                rowNumber++;
            }
            float numFiles = (tempDataTable.GetLength(0) / numEntries);
            int counter; 
            int rows = tempDataTable.GetLength(0);

            for (int x = 0; x < numFiles+1; x++)
            {
               
                if (!Directory.Exists(filePath + "/" + nameNoExt + "/"))
                {
                    Directory.CreateDirectory(filePath + "/" + nameNoExt + "/");
                }

                using (XmlWriter writer = XmlWriter.Create(filePath + "/" + nameNoExt +  "/" + nameNoExt + (x + 1) + ".xml"))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("SetupInvMaster");
                    for (int i = 0; i < numEntries; i++)
                    {
                        counter = i + x * numEntries;

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

                        if (counter+2 > rows)
                        {
                            i = numEntries + 10;
                        }

                    }
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            }
            tempDataTable = null;
            MessageBox.Show("Generation of XML files finished.", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ConvXMLbtn2_Click(object sender, EventArgs e)
        {
            Int32.TryParse(textBox1.Text, out numEntries);

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
            string name = saveFileDialog1.FileName;
            string nameNoExt = Path.GetFileNameWithoutExtension(name);
            filePath = Path.GetDirectoryName(name);

            conn = new OleDbConnection(con);
            oconn = new OleDbCommand("SELECT * FROM ["+worksheetName+"$] WHERE COST<>0", conn);
            conn.Open();
            sda = new OleDbDataAdapter(oconn);
            data = new DataTable();
            sda.Fill(data);
            dataGridView1.DataSource = data;

            string[,] tempDataTable = new string[dataGridView1.Rows.Count, 3];
            int rowNumber = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                tempDataTable[rowNumber, 0] = row.Cells["WHSE"].Value.ToString();
                tempDataTable[rowNumber, 1] = row.Cells["*StockCode"].Value.ToString();
                tempDataTable[rowNumber, 2] = row.Cells["COST"].Value.ToString();
                rowNumber++;
            }

            float numFiles = (tempDataTable.GetLength(0) / numEntries);
            int counter;
            int rows = tempDataTable.GetLength(0);

            for (int x = 0; x < numFiles + 1; x++)
            {
                if (!Directory.Exists(filePath + "/" + nameNoExt + "/"))
                {
                    Directory.CreateDirectory(filePath + "/" + nameNoExt + "/");
                }

                using (XmlWriter writer = XmlWriter.Create(filePath + "/" + nameNoExt + "/" + nameNoExt + (x + 1) + ".xml"))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("PostInvCostChange");
                    for (int i = 0; i < numEntries; i++)
                    {
                        counter = i + x * numEntries;

                        writer.WriteStartElement("Item");
                        writer.WriteElementString("Warehouse", tempDataTable[i, 0]);
                        writer.WriteElementString("StockCode", tempDataTable[i, 1]);
                        writer.WriteElementString("NewUnitCost", tempDataTable[i, 2]);
                        writer.WriteElementString("UpdateLastCost", "Y");
                        writer.WriteElementString("Reference", "UOM CHG");
                        writer.WriteElementString("Notation", "Cost change Note");
                        writer.WriteEndElement();

                        if (counter + 2 > rows)
                        {
                            i = numEntries + 10;
                        }
                    }
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            }
            dataGridView1 = null;
            tempDataTable = null;
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
            }
        }
    }
}

