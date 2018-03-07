namespace 刷新文件卡
{
    partial class FormAddProperty
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.PropertyName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PropertyValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.buttonAddProperty = new System.Windows.Forms.Button();
            this.buttonReturn = new System.Windows.Forms.Button();
            this.openFileDlgAddPro = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.PropertyName,
            this.PropertyValue});
            this.dataGridView1.Location = new System.Drawing.Point(18, 10);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(388, 389);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // PropertyName
            // 
            this.PropertyName.HeaderText = "属性名称";
            this.PropertyName.Name = "PropertyName";
            // 
            // PropertyValue
            // 
            this.PropertyValue.HeaderText = "属性值";
            this.PropertyValue.Name = "PropertyValue";
            // 
            // buttonAddProperty
            // 
            this.buttonAddProperty.Location = new System.Drawing.Point(479, 79);
            this.buttonAddProperty.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.buttonAddProperty.Name = "buttonAddProperty";
            this.buttonAddProperty.Size = new System.Drawing.Size(114, 53);
            this.buttonAddProperty.TabIndex = 1;
            this.buttonAddProperty.Text = "增加属性";
            this.buttonAddProperty.UseVisualStyleBackColor = true;
            this.buttonAddProperty.Click += new System.EventHandler(this.buttonAddProperty_Click);
            // 
            // buttonReturn
            // 
            this.buttonReturn.Location = new System.Drawing.Point(479, 186);
            this.buttonReturn.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.buttonReturn.Name = "buttonReturn";
            this.buttonReturn.Size = new System.Drawing.Size(113, 50);
            this.buttonReturn.TabIndex = 2;
            this.buttonReturn.Text = "返回";
            this.buttonReturn.UseVisualStyleBackColor = true;
            this.buttonReturn.Click += new System.EventHandler(this.buttonReturn_Click);
            // 
            // openFileDlgAddPro
            // 
            this.openFileDlgAddPro.FileName = "openFileDlgAddPro";
            // 
            // FormAddProperty
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(639, 432);
            this.Controls.Add(this.buttonReturn);
            this.Controls.Add(this.buttonAddProperty);
            this.Controls.Add(this.dataGridView1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "FormAddProperty";
            this.Text = "增加属性";
            this.Load += new System.EventHandler(this.FormAddProperty_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn PropertyName;
        private System.Windows.Forms.DataGridViewTextBoxColumn PropertyValue;
        private System.Windows.Forms.Button buttonAddProperty;
        private System.Windows.Forms.Button buttonReturn;
        private System.Windows.Forms.OpenFileDialog openFileDlgAddPro;
    }
}