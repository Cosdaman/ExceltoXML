namespace WindowsFormsApplication2
{
    partial class ReadExcel
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ChooseFileBtn = new System.Windows.Forms.Button();
            this.CloseBtn = new System.Windows.Forms.Button();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.ConvXMLbtn = new System.Windows.Forms.Button();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.ConvXMLbtn2 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // ChooseFileBtn
            // 
            this.ChooseFileBtn.Location = new System.Drawing.Point(12, 12);
            this.ChooseFileBtn.Name = "ChooseFileBtn";
            this.ChooseFileBtn.Size = new System.Drawing.Size(335, 90);
            this.ChooseFileBtn.TabIndex = 0;
            this.ChooseFileBtn.Text = "Choose and Display Excel File";
            this.ChooseFileBtn.UseVisualStyleBackColor = true;
            this.ChooseFileBtn.Click += new System.EventHandler(this.ChooseFileBtn_Click);
            // 
            // CloseBtn
            // 
            this.CloseBtn.Location = new System.Drawing.Point(12, 182);
            this.CloseBtn.Name = "CloseBtn";
            this.CloseBtn.Size = new System.Drawing.Size(242, 53);
            this.CloseBtn.TabIndex = 1;
            this.CloseBtn.Text = "Close Program";
            this.CloseBtn.UseVisualStyleBackColor = true;
            this.CloseBtn.Click += new System.EventHandler(this.CloseBtn_Click);
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(12, 241);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.Height = 33;
            this.dataGridView2.Size = new System.Drawing.Size(1140, 429);
            this.dataGridView2.TabIndex = 2;
            // 
            // ConvXMLbtn
            // 
            this.ConvXMLbtn.Location = new System.Drawing.Point(353, 12);
            this.ConvXMLbtn.Name = "ConvXMLbtn";
            this.ConvXMLbtn.Size = new System.Drawing.Size(335, 90);
            this.ConvXMLbtn.TabIndex = 3;
            this.ConvXMLbtn.Text = "Convert to XML (All)";
            this.ConvXMLbtn.UseVisualStyleBackColor = true;
            this.ConvXMLbtn.Click += new System.EventHandler(this.ConvXMLbtn_Click);
            // 
            // ConvXMLbtn2
            // 
            this.ConvXMLbtn2.Location = new System.Drawing.Point(694, 12);
            this.ConvXMLbtn2.Name = "ConvXMLbtn2";
            this.ConvXMLbtn2.Size = new System.Drawing.Size(335, 90);
            this.ConvXMLbtn2.TabIndex = 4;
            this.ConvXMLbtn2.Text = "Convert to XML (Cost Only)";
            this.ConvXMLbtn2.UseVisualStyleBackColor = true;
            this.ConvXMLbtn2.Click += new System.EventHandler(this.ConvXMLbtn2_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeColumns = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(197, 388);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowTemplate.Height = 33;
            this.dataGridView1.Size = new System.Drawing.Size(116, 21);
            this.dataGridView1.TabIndex = 5;
            this.dataGridView1.Visible = false;
            // 
            // ReadExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1164, 682);
            this.Controls.Add(this.ConvXMLbtn2);
            this.Controls.Add(this.ConvXMLbtn);
            this.Controls.Add(this.CloseBtn);
            this.Controls.Add(this.ChooseFileBtn);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.dataGridView1);
            this.Name = "ReadExcel";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button ChooseFileBtn;
        private System.Windows.Forms.Button CloseBtn;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button ConvXMLbtn;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button ConvXMLbtn2;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}

