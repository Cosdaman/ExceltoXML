using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        DataTable data;

        public ReadExcel()
        {
            InitializeComponent();
        }

        private void ChooseFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePath = file.FileName;
                fileExt = Path.GetExtension(filePath);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        if (fileExt.CompareTo(".xls") == 0)
                            con = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                        else
                            con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';"; //for above excel 2007  

                        conn = new OleDbConnection(con);
                        oconn = new OleDbCommand("SELECT * FROM [Sheet1$]", conn);
                        conn.Open();
                        sda = new OleDbDataAdapter(oconn);
                        DataTable data = new DataTable();
                        sda.Fill(data);
                        dataGridView2.DataSource = data;
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
                tempDataTable[rowNumber, 0] = row.Cells["StockCode"].Value.ToString();
                tempDataTable[rowNumber, 1] = row.Cells["StockUom"].Value.ToString();
                tempDataTable[rowNumber, 2] = row.Cells["AlternateUom"].Value.ToString();
                tempDataTable[rowNumber, 3] = row.Cells["OtherUom"].Value.ToString();
                tempDataTable[rowNumber, 4] = row.Cells["ConvFactAltUom"].Value.ToString();
                tempDataTable[rowNumber, 5] = row.Cells["ConvMulDiv"].Value.ToString();
                tempDataTable[rowNumber, 6] = row.Cells["ConvFactOthUom"].Value.ToString();
                tempDataTable[rowNumber, 7] = row.Cells["MulDiv"].Value.ToString();
                rowNumber++;
            }
            using (XmlWriter writer = XmlWriter.Create(name))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("SetupInvMaster");
                for (int i = 0; i < tempDataTable.GetLength(0); i++)
                {
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
            string name = saveFileDialog1.FileName;

            conn = new OleDbConnection(con);
            oconn = new OleDbCommand("SELECT * FROM [Sheet1$] WHERE cost<>0", conn);
            conn.Open();
            sda = new OleDbDataAdapter(oconn);
            data = new DataTable();
            sda.Fill(data);
            dataGridView1.DataSource = data;

            string[,] tempDataTable = new string[dataGridView1.Rows.Count, 3];
            int rowNumber = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                    tempDataTable[rowNumber, 0] = row.Cells["Warehouse"].Value.ToString();
                    tempDataTable[rowNumber, 1] = row.Cells["StockCode"].Value.ToString();
                    tempDataTable[rowNumber, 2] = row.Cells["Cost"].Value.ToString();
                    rowNumber++;
            }

            using (XmlWriter writer = XmlWriter.Create(name))
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
                writer.WriteEndElement();
                writer.WriteEndDocument();
                MessageBox.Show("Generation of XML file has finished.", "Success!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}

