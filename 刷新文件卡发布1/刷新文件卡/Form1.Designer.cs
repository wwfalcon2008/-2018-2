namespace 刷新文件卡
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button2 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.checkBoxModifyProp = new System.Windows.Forms.CheckBox();
            this.buttonCheck = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.propertyName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.propertyValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.propertyOriginalValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.buttonModifiyProp = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.buttonPrint = new System.Windows.Forms.Button();
            this.checkBoxPDF = new System.Windows.Forms.CheckBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.checkBoxUnmodifyPDFName = new System.Windows.Forms.CheckBox();
            this.buttonSaveAs = new System.Windows.Forms.Button();
            this.buttonSaveAsPDF = new System.Windows.Forms.Button();
            this.buttonClearPrintList = new System.Windows.Forms.Button();
            this.printerSet = new System.Windows.Forms.Button();
            this.printerA4ComboBox = new System.Windows.Forms.ComboBox();
            this.printerA3ComboBox = new System.Windows.Forms.ComboBox();
            this.printerPDFComboBox = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonClose2 = new System.Windows.Forms.Button();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.buttonOrgnization = new System.Windows.Forms.Button();
            this.buttonCreateTable = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.radioButtonAssm = new System.Windows.Forms.RadioButton();
            this.radioButtonPart = new System.Windows.Forms.RadioButton();
            this.buttonRename = new System.Windows.Forms.Button();
            this.buttonRenameRcdClear = new System.Windows.Forms.Button();
            this.textBoxOrgWarning = new System.Windows.Forms.TextBox();
            this.textBoxOrgnization = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.toolTip2 = new System.Windows.Forms.ToolTip(this.components);
            this.textBoxFileListPorp = new System.Windows.Forms.TextBox();
            this.textBoxFileListPrint = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(826, 232);
            this.button1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(140, 80);
            this.button1.TabIndex = 0;
            this.button1.Text = "替换模板\r\n（文件夹）";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.buttonFolder_Click);
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(6, 15);
            this.textBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox1.Size = new System.Drawing.Size(600, 350);
            this.textBox1.TabIndex = 1;
            this.textBox1.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button2.Location = new System.Drawing.Point(826, 142);
            this.button2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(140, 80);
            this.button2.TabIndex = 2;
            this.button2.Text = "替换模板 \r\n（文件）";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.buttonFile_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.button3.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button3.Location = new System.Drawing.Point(826, 592);
            this.button3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(140, 80);
            this.button3.TabIndex = 3;
            this.button3.Text = "关闭程序";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button4.Location = new System.Drawing.Point(826, 52);
            this.button4.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(140, 80);
            this.button4.TabIndex = 4;
            this.button4.Text = "清空记录";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 723);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 15);
            this.label1.TabIndex = 5;
            // 
            // checkBoxModifyProp
            // 
            this.checkBoxModifyProp.AutoSize = true;
            this.checkBoxModifyProp.Location = new System.Drawing.Point(887, 17);
            this.checkBoxModifyProp.Margin = new System.Windows.Forms.Padding(4);
            this.checkBoxModifyProp.Name = "checkBoxModifyProp";
            this.checkBoxModifyProp.Size = new System.Drawing.Size(89, 19);
            this.checkBoxModifyProp.TabIndex = 8;
            this.checkBoxModifyProp.Text = "移植属性";
            this.checkBoxModifyProp.UseVisualStyleBackColor = true;
            // 
            // buttonCheck
            // 
            this.buttonCheck.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.buttonCheck.Location = new System.Drawing.Point(826, 322);
            this.buttonCheck.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonCheck.Name = "buttonCheck";
            this.buttonCheck.Size = new System.Drawing.Size(140, 80);
            this.buttonCheck.TabIndex = 9;
            this.buttonCheck.Text = "检测空属性";
            this.buttonCheck.UseVisualStyleBackColor = true;
            this.buttonCheck.Click += new System.EventHandler(this.buttonCheck_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(6, 369);
            this.textBox2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.textBox2.Multiline = true;
            this.textBox2.Name = "textBox2";
            this.textBox2.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox2.Size = new System.Drawing.Size(800, 150);
            this.textBox2.TabIndex = 12;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.propertyName,
            this.propertyValue,
            this.propertyOriginalValue});
            this.dataGridView1.Location = new System.Drawing.Point(6, 523);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(800, 120);
            this.dataGridView1.TabIndex = 13;
            // 
            // propertyName
            // 
            this.propertyName.DividerWidth = 2;
            this.propertyName.HeaderText = "属性名称";
            this.propertyName.MinimumWidth = 100;
            this.propertyName.Name = "propertyName";
            this.propertyName.Width = 180;
            // 
            // propertyValue
            // 
            this.propertyValue.DividerWidth = 2;
            this.propertyValue.HeaderText = "属性内容";
            this.propertyValue.MinimumWidth = 100;
            this.propertyValue.Name = "propertyValue";
            this.propertyValue.Width = 180;
            // 
            // propertyOriginalValue
            // 
            this.propertyOriginalValue.DividerWidth = 2;
            this.propertyOriginalValue.HeaderText = "属性原值";
            this.propertyOriginalValue.MinimumWidth = 100;
            this.propertyOriginalValue.Name = "propertyOriginalValue";
            this.propertyOriginalValue.Width = 180;
            // 
            // buttonModifiyProp
            // 
            this.buttonModifiyProp.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.buttonModifiyProp.Location = new System.Drawing.Point(826, 502);
            this.buttonModifiyProp.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonModifiyProp.Name = "buttonModifiyProp";
            this.buttonModifiyProp.Size = new System.Drawing.Size(140, 80);
            this.buttonModifiyProp.TabIndex = 14;
            this.buttonModifiyProp.Text = "编辑属性";
            this.buttonModifiyProp.UseVisualStyleBackColor = true;
            this.buttonModifiyProp.Click += new System.EventHandler(this.buttonModifiyProp_Click);
            // 
            // button5
            // 
            this.button5.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.button5.Location = new System.Drawing.Point(826, 412);
            this.button5.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(140, 80);
            this.button5.TabIndex = 15;
            this.button5.Text = "清空属性\r\n修改清单";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // buttonPrint
            // 
            this.buttonPrint.Location = new System.Drawing.Point(826, 322);
            this.buttonPrint.Name = "buttonPrint";
            this.buttonPrint.Size = new System.Drawing.Size(140, 80);
            this.buttonPrint.TabIndex = 16;
            this.buttonPrint.Text = "打印";
            this.buttonPrint.UseVisualStyleBackColor = true;
            this.buttonPrint.Click += new System.EventHandler(this.buttonPrint_Click);
            // 
            // checkBoxPDF
            // 
            this.checkBoxPDF.AutoSize = true;
            this.checkBoxPDF.Location = new System.Drawing.Point(826, 91);
            this.checkBoxPDF.Name = "checkBoxPDF";
            this.checkBoxPDF.Size = new System.Drawing.Size(98, 19);
            this.checkBoxPDF.TabIndex = 17;
            this.checkBoxPDF.Text = "打印到PDF";
            this.checkBoxPDF.UseVisualStyleBackColor = true;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(991, 708);
            this.tabControl1.TabIndex = 18;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.textBoxFileListPorp);
            this.tabPage1.Controls.Add(this.textBox1);
            this.tabPage1.Controls.Add(this.textBox2);
            this.tabPage1.Controls.Add(this.dataGridView1);
            this.tabPage1.Controls.Add(this.button5);
            this.tabPage1.Controls.Add(this.checkBoxModifyProp);
            this.tabPage1.Controls.Add(this.buttonModifiyProp);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.buttonCheck);
            this.tabPage1.Controls.Add(this.button2);
            this.tabPage1.Controls.Add(this.button3);
            this.tabPage1.Controls.Add(this.button4);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(983, 679);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "属性操作";
            this.tabPage1.UseVisualStyleBackColor = true;
            this.tabPage1.Click += new System.EventHandler(this.tabPage1_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.textBoxFileListPrint);
            this.tabPage2.Controls.Add(this.checkBoxUnmodifyPDFName);
            this.tabPage2.Controls.Add(this.buttonSaveAs);
            this.tabPage2.Controls.Add(this.buttonSaveAsPDF);
            this.tabPage2.Controls.Add(this.buttonClearPrintList);
            this.tabPage2.Controls.Add(this.printerSet);
            this.tabPage2.Controls.Add(this.printerA4ComboBox);
            this.tabPage2.Controls.Add(this.printerA3ComboBox);
            this.tabPage2.Controls.Add(this.printerPDFComboBox);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.buttonClose2);
            this.tabPage2.Controls.Add(this.textBox4);
            this.tabPage2.Controls.Add(this.textBox3);
            this.tabPage2.Controls.Add(this.buttonPrint);
            this.tabPage2.Controls.Add(this.checkBoxPDF);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(983, 679);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "打印输出";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // checkBoxUnmodifyPDFName
            // 
            this.checkBoxUnmodifyPDFName.AutoSize = true;
            this.checkBoxUnmodifyPDFName.Location = new System.Drawing.Point(826, 117);
            this.checkBoxUnmodifyPDFName.Name = "checkBoxUnmodifyPDFName";
            this.checkBoxUnmodifyPDFName.Size = new System.Drawing.Size(128, 19);
            this.checkBoxUnmodifyPDFName.TabIndex = 28;
            this.checkBoxUnmodifyPDFName.Text = "不修改PDF名称";
            this.checkBoxUnmodifyPDFName.UseVisualStyleBackColor = true;
            // 
            // buttonSaveAs
            // 
            this.buttonSaveAs.Location = new System.Drawing.Point(826, 142);
            this.buttonSaveAs.Name = "buttonSaveAs";
            this.buttonSaveAs.Size = new System.Drawing.Size(140, 80);
            this.buttonSaveAs.TabIndex = 27;
            this.buttonSaveAs.Text = "另存到STP";
            this.buttonSaveAs.UseVisualStyleBackColor = true;
            this.buttonSaveAs.Click += new System.EventHandler(this.buttonSaveAs_Click);
            // 
            // buttonSaveAsPDF
            // 
            this.buttonSaveAsPDF.Location = new System.Drawing.Point(826, 232);
            this.buttonSaveAsPDF.Name = "buttonSaveAsPDF";
            this.buttonSaveAsPDF.Size = new System.Drawing.Size(140, 80);
            this.buttonSaveAsPDF.TabIndex = 26;
            this.buttonSaveAsPDF.Text = "另存为PDF";
            this.buttonSaveAsPDF.UseVisualStyleBackColor = true;
            this.buttonSaveAsPDF.Click += new System.EventHandler(this.buttonSaveAsPDF_Click);
            // 
            // buttonClearPrintList
            // 
            this.buttonClearPrintList.Location = new System.Drawing.Point(826, 502);
            this.buttonClearPrintList.Name = "buttonClearPrintList";
            this.buttonClearPrintList.Size = new System.Drawing.Size(140, 80);
            this.buttonClearPrintList.TabIndex = 24;
            this.buttonClearPrintList.Text = "清空打印清单";
            this.buttonClearPrintList.UseVisualStyleBackColor = true;
            this.buttonClearPrintList.Click += new System.EventHandler(this.buttonClearPrintList_Click);
            // 
            // printerSet
            // 
            this.printerSet.Location = new System.Drawing.Point(826, 412);
            this.printerSet.Name = "printerSet";
            this.printerSet.Size = new System.Drawing.Size(140, 80);
            this.printerSet.TabIndex = 23;
            this.printerSet.Text = "选定打印机";
            this.printerSet.UseVisualStyleBackColor = true;
            this.printerSet.Click += new System.EventHandler(this.printerSet_Click);
            // 
            // printerA4ComboBox
            // 
            this.printerA4ComboBox.FormattingEnabled = true;
            this.printerA4ComboBox.Location = new System.Drawing.Point(133, 596);
            this.printerA4ComboBox.Name = "printerA4ComboBox";
            this.printerA4ComboBox.Size = new System.Drawing.Size(612, 23);
            this.printerA4ComboBox.TabIndex = 22;
            // 
            // printerA3ComboBox
            // 
            this.printerA3ComboBox.FormattingEnabled = true;
            this.printerA3ComboBox.Location = new System.Drawing.Point(133, 561);
            this.printerA3ComboBox.Name = "printerA3ComboBox";
            this.printerA3ComboBox.Size = new System.Drawing.Size(612, 23);
            this.printerA3ComboBox.TabIndex = 22;
            // 
            // printerPDFComboBox
            // 
            this.printerPDFComboBox.FormattingEnabled = true;
            this.printerPDFComboBox.Location = new System.Drawing.Point(133, 531);
            this.printerPDFComboBox.Name = "printerPDFComboBox";
            this.printerPDFComboBox.Size = new System.Drawing.Size(612, 23);
            this.printerPDFComboBox.TabIndex = 22;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(21, 599);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(98, 15);
            this.label4.TabIndex = 21;
            this.label4.Text = "选择A4打印机";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 564);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(98, 15);
            this.label3.TabIndex = 21;
            this.label3.Text = "选择A3打印机";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(21, 534);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(106, 15);
            this.label2.TabIndex = 21;
            this.label2.Text = "选择PDF打印机";
            // 
            // buttonClose2
            // 
            this.buttonClose2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.buttonClose2.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.buttonClose2.Location = new System.Drawing.Point(826, 592);
            this.buttonClose2.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.buttonClose2.Name = "buttonClose2";
            this.buttonClose2.Size = new System.Drawing.Size(140, 80);
            this.buttonClose2.TabIndex = 20;
            this.buttonClose2.Text = "关闭程序";
            this.buttonClose2.UseVisualStyleBackColor = false;
            this.buttonClose2.Click += new System.EventHandler(this.buttonClose2_Click);
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(6, 371);
            this.textBox4.Multiline = true;
            this.textBox4.Name = "textBox4";
            this.textBox4.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox4.Size = new System.Drawing.Size(800, 150);
            this.textBox4.TabIndex = 19;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(6, 15);
            this.textBox3.Multiline = true;
            this.textBox3.Name = "textBox3";
            this.textBox3.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBox3.Size = new System.Drawing.Size(600, 350);
            this.textBox3.TabIndex = 18;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.buttonOrgnization);
            this.tabPage3.Controls.Add(this.buttonCreateTable);
            this.tabPage3.Controls.Add(this.button7);
            this.tabPage3.Controls.Add(this.button6);
            this.tabPage3.Controls.Add(this.radioButtonAssm);
            this.tabPage3.Controls.Add(this.radioButtonPart);
            this.tabPage3.Controls.Add(this.buttonRename);
            this.tabPage3.Controls.Add(this.buttonRenameRcdClear);
            this.tabPage3.Controls.Add(this.textBoxOrgWarning);
            this.tabPage3.Controls.Add(this.textBoxOrgnization);
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(983, 679);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "文件名整理";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // buttonOrgnization
            // 
            this.buttonOrgnization.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.buttonOrgnization.Location = new System.Drawing.Point(826, 350);
            this.buttonOrgnization.Name = "buttonOrgnization";
            this.buttonOrgnization.Size = new System.Drawing.Size(140, 80);
            this.buttonOrgnization.TabIndex = 9;
            this.buttonOrgnization.Text = "整理文件名";
            this.buttonOrgnization.UseVisualStyleBackColor = true;
            this.buttonOrgnization.Click += new System.EventHandler(this.buttonOrgnization_Click);
            // 
            // buttonCreateTable
            // 
            this.buttonCreateTable.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.buttonCreateTable.Location = new System.Drawing.Point(826, 260);
            this.buttonCreateTable.Name = "buttonCreateTable";
            this.buttonCreateTable.Size = new System.Drawing.Size(140, 80);
            this.buttonCreateTable.TabIndex = 8;
            this.buttonCreateTable.Text = "生成表格";
            this.buttonCreateTable.UseVisualStyleBackColor = true;
            this.buttonCreateTable.Click += new System.EventHandler(this.buttonCreateTable_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(843, 491);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(47, 30);
            this.button7.TabIndex = 7;
            this.button7.Text = "button7";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button6
            // 
            this.button6.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.button6.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.button6.Location = new System.Drawing.Point(826, 592);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(140, 80);
            this.button6.TabIndex = 6;
            this.button6.Text = "关闭程序";
            this.button6.UseVisualStyleBackColor = false;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // radioButtonAssm
            // 
            this.radioButtonAssm.AutoSize = true;
            this.radioButtonAssm.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.radioButtonAssm.Location = new System.Drawing.Point(826, 40);
            this.radioButtonAssm.Name = "radioButtonAssm";
            this.radioButtonAssm.Size = new System.Drawing.Size(148, 24);
            this.radioButtonAssm.TabIndex = 5;
            this.radioButtonAssm.Text = "装配体(SLDASM)";
            this.radioButtonAssm.UseVisualStyleBackColor = true;
            // 
            // radioButtonPart
            // 
            this.radioButtonPart.AutoSize = true;
            this.radioButtonPart.Checked = true;
            this.radioButtonPart.Font = new System.Drawing.Font("微软雅黑", 9F);
            this.radioButtonPart.Location = new System.Drawing.Point(826, 15);
            this.radioButtonPart.Name = "radioButtonPart";
            this.radioButtonPart.Size = new System.Drawing.Size(126, 24);
            this.radioButtonPart.TabIndex = 4;
            this.radioButtonPart.TabStop = true;
            this.radioButtonPart.Text = "零件(SLDPRT)";
            this.radioButtonPart.UseVisualStyleBackColor = true;
            // 
            // buttonRename
            // 
            this.buttonRename.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.buttonRename.Location = new System.Drawing.Point(826, 170);
            this.buttonRename.Name = "buttonRename";
            this.buttonRename.Size = new System.Drawing.Size(140, 80);
            this.buttonRename.TabIndex = 3;
            this.buttonRename.Text = "文件名整理";
            this.buttonRename.UseVisualStyleBackColor = true;
            this.buttonRename.Click += new System.EventHandler(this.buttonRename_Click);
            // 
            // buttonRenameRcdClear
            // 
            this.buttonRenameRcdClear.Font = new System.Drawing.Font("微软雅黑", 12F);
            this.buttonRenameRcdClear.Location = new System.Drawing.Point(826, 80);
            this.buttonRenameRcdClear.Name = "buttonRenameRcdClear";
            this.buttonRenameRcdClear.Size = new System.Drawing.Size(140, 80);
            this.buttonRenameRcdClear.TabIndex = 2;
            this.buttonRenameRcdClear.Text = "清空记录";
            this.buttonRenameRcdClear.UseVisualStyleBackColor = true;
            this.buttonRenameRcdClear.Click += new System.EventHandler(this.buttonRenameRcdClear_Click);
            // 
            // textBoxOrgWarning
            // 
            this.textBoxOrgWarning.Location = new System.Drawing.Point(6, 371);
            this.textBoxOrgWarning.Multiline = true;
            this.textBoxOrgWarning.Name = "textBoxOrgWarning";
            this.textBoxOrgWarning.Size = new System.Drawing.Size(800, 150);
            this.textBoxOrgWarning.TabIndex = 1;
            // 
            // textBoxOrgnization
            // 
            this.textBoxOrgnization.Location = new System.Drawing.Point(6, 15);
            this.textBoxOrgnization.Multiline = true;
            this.textBoxOrgnization.Name = "textBoxOrgnization";
            this.textBoxOrgnization.Size = new System.Drawing.Size(800, 350);
            this.textBoxOrgnization.TabIndex = 0;
            // 
            // toolTip1
            // 
            this.toolTip1.Popup += new System.Windows.Forms.PopupEventHandler(this.toolTip1_Popup);
            // 
            // textBoxFileListPorp
            // 
            this.textBoxFileListPorp.Location = new System.Drawing.Point(620, 15);
            this.textBoxFileListPorp.Multiline = true;
            this.textBoxFileListPorp.Name = "textBoxFileListPorp";
            this.textBoxFileListPorp.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxFileListPorp.Size = new System.Drawing.Size(180, 350);
            this.textBoxFileListPorp.TabIndex = 16;
            // 
            // textBoxFileListPrint
            // 
            this.textBoxFileListPrint.Location = new System.Drawing.Point(620, 15);
            this.textBoxFileListPrint.Multiline = true;
            this.textBoxFileListPrint.Name = "textBoxFileListPrint";
            this.textBoxFileListPrint.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.textBoxFileListPrint.Size = new System.Drawing.Size(180, 350);
            this.textBoxFileListPrint.TabIndex = 29;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1006, 743);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "Form1";
            this.Text = "SolidWorks辅助工具2018";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBoxModifyProp;
        private System.Windows.Forms.Button buttonCheck;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button buttonModifiyProp;
        private System.Windows.Forms.DataGridViewTextBoxColumn propertyName;
        private System.Windows.Forms.DataGridViewTextBoxColumn propertyValue;
        private System.Windows.Forms.DataGridViewTextBoxColumn propertyOriginalValue;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button buttonPrint;
        private System.Windows.Forms.CheckBox checkBoxPDF;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Button buttonClose2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox printerPDFComboBox;
        private System.Windows.Forms.ComboBox printerA4ComboBox;
        private System.Windows.Forms.ComboBox printerA3ComboBox;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button printerSet;
        private System.Windows.Forms.Button buttonClearPrintList;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button buttonRename;
        private System.Windows.Forms.Button buttonRenameRcdClear;
        private System.Windows.Forms.TextBox textBoxOrgWarning;
        private System.Windows.Forms.TextBox textBoxOrgnization;
        private System.Windows.Forms.RadioButton radioButtonAssm;
        private System.Windows.Forms.RadioButton radioButtonPart;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button buttonSaveAsPDF;
        private System.Windows.Forms.Button buttonSaveAs;
        private System.Windows.Forms.CheckBox checkBoxUnmodifyPDFName;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ToolTip toolTip2;
        private System.Windows.Forms.Button buttonCreateTable;
        private System.Windows.Forms.Button buttonOrgnization;
        private System.Windows.Forms.TextBox textBoxFileListPorp;
        private System.Windows.Forms.TextBox textBoxFileListPrint;
    }
}

