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
            this.label1 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
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
            this.CloseBtn.Location = new System.Drawing.Point(12, 239);
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
            this.dataGridView2.Location = new System.Drawing.Point(12, 298);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.RowTemplate.Height = 33;
            this.dataGridView2.Size = new System.Drawing.Size(1946, 382);
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
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.label1.Location = new System.Drawing.Point(12, 128);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(447, 37);
            this.label1.TabIndex = 6;
            this.label1.Text = "Enter # of entries per XML file:";
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.textBox1.Location = new System.Drawing.Point(465, 125);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(152, 44);
            this.textBox1.TabIndex = 7;
            this.textBox1.Text = "500";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(694, 108);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(335, 59);
            this.button1.TabIndex = 8;
            this.button1.Text = "Open Directory";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.label2.Location = new System.Drawing.Point(1035, 12);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(726, 37);
            this.label2.TabIndex = 9;
            this.label2.Text = "Ensure that there are no duplicate column names.";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.label3.Location = new System.Drawing.Point(1035, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(750, 37);
            this.label3.TabIndex = 10;
            this.label3.Text = "First row is considered as the name of the columns.";
            // 
            // ReadExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1970, 692);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ConvXMLbtn2);
            this.Controls.Add(this.ConvXMLbtn);
            this.Controls.Add(this.CloseBtn);
            this.Controls.Add(this.ChooseFileBtn);
            this.Controls.Add(this.dataGridView2);
            this.Name = "ReadExcel";
            this.Text = "Form1";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ChooseFileBtn;
        private System.Windows.Forms.Button CloseBtn;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.Button ConvXMLbtn;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button ConvXMLbtn2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
    }
}

