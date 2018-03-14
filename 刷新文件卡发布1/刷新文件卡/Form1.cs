using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using SldWorks;
using SwCommands;
using SwConst;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


using System.Drawing.Printing;
//using Component = SolidWorks.Interop.sldworks.Component;
//using View = SolidWorks.Interop.sldworks.View;
//using swCustomInfoType_e = SolidWorks.Interop.swconst.swCustomInfoType_e;
//using swCustomPropertyAddOption_e = SolidWorks.Interop.swconst.swCustomPropertyAddOption_e;
//using swDwgPaperSizes_e = SolidWorks.Interop.swconst.swDwgPaperSizes_e;
//using CustomPropertyManager = SolidWorks.Interop.sldworks.CustomPropertyManager;
//using DrawingDoc = SolidWorks.Interop.sldworks.DrawingDoc;
//using ModelDoc2 = SolidWorks.Interop.sldworks.ModelDoc2;
//using ModelDocExtension = SolidWorks.Interop.sldworks.ModelDocExtension;
//using Sheet = SolidWorks.Interop.sldworks.Sheet;
//using PageSetup = SolidWorks.Interop.sldworks.PageSetup;
//using PrintSpecification = SolidWorks.Interop.sldworks.PrintSpecification;


// master branch changed in csharp file 


namespace 刷新文件卡
{
    public partial class Form1 : Form
    {
        ExcelEdit excel;
        int row;
        string sheetName="";
        string path = "";// = @"D:\SheetTemplate\";
        string xlsName;
        string valPartNumber = "";
        string valoutPartNumber = "";
        bool statusPartNumber;
        string valPartName = "";
        string valoutPartName = "";
        bool statusPartName;
        object[] configureNames;
        //TextBox textBox1 = new TextBox();
        static public string AUTHOR = "北京海光仪器有限公司";
        public Form1()
        {
            InitializeComponent();
            label1.Text = AUTHOR;

            //针对checkbox：不修改pdf文件名称的弹出式提示
            toolTip1.SetToolTip(checkBoxUnmodifyPDFName, "用于将无对应模型（SLDPRT或SLDASM）的工程图另存为PDF文件");
            toolTip2.SetToolTip(checkBoxPDF, "用于设定打印机为pdf虚拟打印机，在打印操作中不打印纸质文件");
        }
        int IErrors = 0;
        int IWarnings = 0;
        string sheetformatPath = @"D:\SheetTemplate\";
        double[] sheetProperties = null;
        //double sheetScale = 0;
        SwConst.swDwgPaperSizes_e paperSize;
        double width = 0;
        double height = 0;
        string filePath = "";
        string[] fileNames = null;
        int counter = 0;
        int total = 0;
        string printerPDF = "";
        string printerA4 = "";
        string printerA3 = "";
        //string defaultPrinter = "";
        PrintDocument print = new PrintDocument();


        private void buttonFolder_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this,@"将目标模板全部复制到D:\SheetTemplate\并按确定来选择图纸存放的文件夹");
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                path = folderBrowserDialog1.SelectedPath;
                DirectoryInfo dir = new DirectoryInfo(path);
                FileInfo[] inf = dir.GetFiles();
                total = 0;

                foreach (var fileInfo in inf)
                {
                    if (fileInfo.Extension.ToUpper().Equals(".SLDDRW"))
                    {
                        total++;
                    }
                }

                countDown();
                SldWorks.SldWorks swApp = new SldWorks.SldWorks();
                counter = 0;

                foreach (FileInfo f in inf)
                {
                    if (f.Extension.ToUpper().Equals(".SLDDRW") && (f.Name[0] != '~'))
                    {
                        filePath = path + '\\' + f.ToString();

                        ModelDoc2 swDoc = swApp.OpenDoc6(filePath,
                            3, 1, null, ref IErrors, ref IWarnings);


                        DrawingDoc swDrwng = default(DrawingDoc);
                        swDrwng = (DrawingDoc)swDoc;
                        Sheet activeSheet = default(Sheet);
                        activeSheet = (Sheet)swDrwng.GetCurrentSheet();
                        Debug.Print("Active sheet name: " + activeSheet.GetName());
                        string templatePath = null;
                        string templateName = null;

                        templatePath = activeSheet.GetTemplateName();
                        string[] tempStrings = templatePath.Split('\\');
                        templateName = sheetformatPath + tempStrings[tempStrings.Length - 1];
                        Debug.Print("Sheet format template name to modify: " + templateName);
                        activeSheet.SetTemplateName(templateName);
                        sheetProperties = (double[])activeSheet.GetProperties();
                        paperSize = (swDwgPaperSizes_e)activeSheet.GetSize(ref width, ref height);
                        sheetName = activeSheet.GetName();

                        swDrwng.SetupSheet5(sheetName, (short)paperSize, (short)sheetProperties[1],
                            (double)sheetProperties[2], (double)sheetProperties[3], true, templateName, width, height,
                            "默认", true);
                        activeSheet.ReloadTemplate(false);
                        if (checkBoxModifyProp.Checked == true)
                        {
                            modifyProperty();
                        }


                        swDoc.Visible = true;
                        swDoc.Save();
                        counter++;
                        textBox1.AppendText(counter + ": " + filePath + System.Environment.NewLine);
                        textBoxFileListPorp.AppendText("\"" + filePath + "\" ");

                        try
                        {
                            swApp.CloseDoc(swDoc.GetPathName());
                        }
                        catch (System.ComponentModel.Win32Exception we)
                        {

                            MessageBox.Show(this, we.Message);
                            return;
                            throw;
                        }

                    }
                }
                textBoxFileListPorp.AppendText("-------------" + System.Environment.NewLine);

                MessageBox.Show(this,"替换完成,共替换" + counter + "个文件");
                counter = 0;
                total = 0;
            }

        }

        private void buttonFile_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this,@"将目标模板全部复制到D:\SheetTemplate\并按确定来选择图纸存放的文件夹");
            openFileDialog1.Filter = "SLDDRW files (*.slddrw)|*.slddrw";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNames = openFileDialog1.FileNames;
                total = fileNames.Length;
                countDown();
                SldWorks.SldWorks swApp = new SldWorks.SldWorks();

                foreach (string fileName in fileNames)
                {
                    string[] test = fileName.Split('.');

                    if (test[test.Length - 1].ToUpper().Equals("SLDDRW") && (fileName[0] != '~'))
                    {
                        filePath = fileName;

                        ModelDoc2 swDoc = swApp.OpenDoc6(filePath,
                            3, 1, null, ref IErrors, ref IWarnings);

                        Debug.WriteIf(swDoc == null, swDoc.ToString() + " IS NULL");
                        DrawingDoc swDrwng = default(DrawingDoc);


                        swDrwng = (DrawingDoc)swDoc;
                        Debug.WriteIf(swDrwng == null, swDoc.ToString() + "  " + swDrwng.ToString());
                        Sheet activeSheet = default(Sheet);
                        activeSheet = (Sheet)swDrwng.GetCurrentSheet();
                        Debug.Print("Active sheet name: " + activeSheet.GetName());
                        string templatePath = null;
                        string templateName = null;

                        templatePath = activeSheet.GetTemplateName();
                        string[] tempStrings = templatePath.Split('\\');
                        templateName = sheetformatPath + tempStrings[tempStrings.Length - 1];
                        Debug.Print("Sheet format template name to modify: " + templateName);
                        activeSheet.SetTemplateName(templateName);
                        sheetProperties = (double[])activeSheet.GetProperties();
                        paperSize = (swDwgPaperSizes_e)activeSheet.GetSize(ref width, ref height);


                        sheetName = activeSheet.GetName();

                        swDrwng.SetupSheet5(sheetName, (short)paperSize, (short)sheetProperties[1],
                            (double)sheetProperties[2], (double)sheetProperties[3], true, templateName, width, height,
                            "默认", true);
                        activeSheet.ReloadTemplate(false);
                        Debug.Print("File name to modify: " + activeSheet.GetName());
                        if (checkBoxModifyProp.Checked == true)
                        {
                            modifyProperty();
                        }
                        swDoc.Visible = true;
                        swDoc.Save();
                        counter++;
                        textBox1.AppendText(counter + ": " + filePath + System.Environment.NewLine);
                        textBoxFileListPorp.AppendText("\"" + filePath + "\" ");
                        try
                        {
                            swApp.CloseDoc(swDoc.GetPathName());
                        }
                        catch (System.ComponentModel.Win32Exception we)
                        {

                            MessageBox.Show(this, we.Message);
                            return;
                            throw;
                        }
                    }
                }
                textBoxFileListPorp.AppendText("-------------" + System.Environment.NewLine);

                MessageBox.Show(this,"替换完成,共替换" + counter + "个文件");
                counter = 0;
                total = 0;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(this,"确认清空记录栏？", "提示", MessageBoxButtons.OKCancel);
            if (dr == DialogResult.OK)
            {
                textBox1.Text = "";
                textBox2.Text = "";
                textBoxFileListPorp.Text = "";
                label1.Text = AUTHOR;
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            countDown();
            textBox1.ScrollToCaret();

        }

        private void countDown()
        {
            label1.Text = "共有" + total + "张图纸要处理，尚余" + (total - counter) + "张。";

        }

        private void modifyProperty()
        {
            SldWorks.SldWorks swApp = new SldWorks.SldWorks();

            ModelDoc2 swDoc = swApp.ActiveDoc;

            object[] swPart = swDoc.GetDependencies2(true, true, true);
            foreach (var o in swPart)
            {
                Debug.Print((string)o);
            }

            //////////////////////////需要判断是PART还是ASSEMBLY///////////
            ModelDoc2 swDocDepend = swApp.GetOpenDocumentByName((string)swPart[1]);
            if (swDocDepend != null)
            {

                Debug.Print("swPart[1]: " + (string)swPart[1]);
                Debug.Print("IErrors: " + IErrors);
                Debug.Print("IWarnings: " + IWarnings);

                /////////// MODIFY CustomProperty////////

                #region Get Properties

                ModelDocExtension swDocExtension = default(ModelDocExtension);
                swDocExtension = swDocDepend.Extension;
                CustomPropertyManager swDocProp = default(CustomPropertyManager);

                var configureNames = (object[])swDocDepend.GetConfigurationNames();
                swDocProp = swDocExtension.CustomPropertyManager[(string)configureNames[0]];

                if (swDocProp != null)
                {
                    string valPartNumber = "";
                    string valoutPartNumber = "";
                    string valCompNumber = "";
                    string valoutCompNumber = "";
                    // 图样名称，规格，表面涂饰，零件类型，热处理，备注，物料编码，备用1，备用2，备用3
                    // 材料？ 部件名称
                    string valPartName = "";
                    string valoutPartName = "";
                    string valTypeSpec = "";
                    string valoutTypeSpec = "";
                    string valSurPro = "";
                    string valoutSurPro = "";
                    string valPartType = "";
                    string valoutPartType = "";
                    string valHeatTreat = "";
                    string valoutHeatTreat = "";
                    string valTechComment = "";
                    string valoutTechComment = "";
                    string valItCode = "";
                    string valoutItCode = "";
                    string valBackup1 = "";
                    string valoutBackup1 = "";
                    string valBackup2 = "";
                    string valoutBackup2 = "";
                    string valBackup3 = "";
                    string valoutBackup3 = "";
                    string valMaterial = "";
                    string valoutMaterial = "";
                    string valCompName = "";
                    string valoutCompName = "";


                    var statusPartNumber = swDocProp.Get4("图样代号", false, out valPartNumber, out valoutPartNumber);
                    var statusCompNumber = swDocProp.Get4("部件代号", false, out valCompNumber, out valoutCompNumber);
                    /////次要内容
                    var statusPartName = swDocProp.Get4("图样名称", false, out valPartName, out valoutPartName);
                    var statusTypeSpec = swDocProp.Get4("规格", false, out valTypeSpec, out valoutTypeSpec);
                    var statusSurPro = swDocProp.Get4("表面涂饰", false, out valSurPro, out valoutSurPro);
                    var statusPartType = swDocProp.Get4("零件类型", false, out valPartType, out valoutPartType);
                    var statusHeatTreat = swDocProp.Get4("热处理", false, out valHeatTreat, out valoutHeatTreat);
                    var statusTechComment = swDocProp.Get4("备注", false, out valTechComment, out valoutTechComment);
                    var statusItCode = swDocProp.Get4("物料编码", false, out valItCode, out valoutItCode);
                    var statusBackup1 = swDocProp.Get4("备用1", false, out valBackup1, out valoutBackup1);
                    var statusBackup2 = swDocProp.Get4("备用1", false, out valBackup2, out valoutBackup2);
                    var statusBackup3 = swDocProp.Get4("备用1", false, out valBackup3, out valoutBackup3);
                    var statusMaterial = swDocProp.Get4("材料", false, out valMaterial, out valoutMaterial);
                    var statusCompName = swDocProp.Get4("部件名称", false, out valCompName, out valoutCompName);



                    Debug.Print("valoutPartNubmer: " + valoutPartNumber);
                    Debug.Print("valPartNubmer: " + valPartNumber);
                    Debug.Print("statusPartNumber:  " + statusPartNumber);
                    Debug.Print("valoutCompNumber:  " + valoutCompNumber);
                    Debug.Print("valCompNubmer: " + valCompNumber);
                    Debug.Print("statusCompNumber:  " + statusCompNumber);
                    if (valoutCompNumber == "")
                    {
                        textBox1.AppendText("WARNING（部件代号为空）: " + filePath + System.Environment.NewLine);
                        textBox2.AppendText("WARNING（部件代号为空）: " + filePath + System.Environment.NewLine);
                    }
                    if (valoutPartNumber == "")
                    {
                        textBox1.AppendText("WARNING（图样代号为空）: " + filePath + System.Environment.NewLine);
                        textBox2.AppendText("WARNING（图样代号为空）: " + filePath + System.Environment.NewLine);
                    }

                    #endregion

                    #region Custom Properties

                    var cusPropMgr = swDoc.Extension.get_CustomPropertyManager("");

                    cusPropMgr.Add3("图样代号", (int)swCustomInfoType_e.swCustomInfoText, valoutPartNumber,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("部件代号", (int)swCustomInfoType_e.swCustomInfoText, valoutCompNumber,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    /////次要内容
                    cusPropMgr.Add3("图样名称", (int)swCustomInfoType_e.swCustomInfoText, valoutPartName,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("规格", (int)swCustomInfoType_e.swCustomInfoText, valoutSurPro,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("表面涂饰", (int)swCustomInfoType_e.swCustomInfoText, valoutSurPro,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("零件类型", (int)swCustomInfoType_e.swCustomInfoText, valoutPartType,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("热处理", (int)swCustomInfoType_e.swCustomInfoText, valoutHeatTreat,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("备注", (int)swCustomInfoType_e.swCustomInfoText, valoutTechComment,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("物料编码", (int)swCustomInfoType_e.swCustomInfoText, valoutItCode,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("备用1", (int)swCustomInfoType_e.swCustomInfoText, valoutBackup1,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("备用2", (int)swCustomInfoType_e.swCustomInfoText, valoutBackup2,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("备用3", (int)swCustomInfoType_e.swCustomInfoText, valoutBackup3,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("部件名称", (int)swCustomInfoType_e.swCustomInfoText, valoutCompName,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                    cusPropMgr.Add3("材料", (int)swCustomInfoType_e.swCustomInfoText, valoutMaterial,
                        (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);

                    #endregion
                }
                else
                {
                    textBox1.AppendText("Warning(swDocProp==null): " + (string)swDoc.GetTitle());
                }
            }
            else
            {
                textBox2.AppendText("找不到参考文件" + (string)swPart[1]);
            }
        }


        private void buttonCheck_Click(object sender, EventArgs e)
        {

            openFileDialog1.Filter = "SLDPRT, SLDASM files (*.SLDPRT,*.SLDASM)|*.sldprt;*.sldasm";
            openFileDialog1.Multiselect = true;
            counter = 0;
            int counterError = 0;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNames = openFileDialog1.FileNames;
                total = fileNames.Length;
                countDown();
                SldWorks.SldWorks swApp = new SldWorks.SldWorks();

                foreach (string fileName in fileNames)
                {
                    string[] test = fileName.Split('.');
                    filePath = fileName;
                    int fileType = 0;
                    counter++;
                    if ((test[test.Length - 1].ToUpper().Equals("SLDASM") || test[test.Length - 1].ToUpper().Equals("SLDPRT")) && fileName[0] != '~')
                    {
                        //(test[test.Length - 1].Equals("SLDASM") || test[test.Length - 1].Equals("SLDPRT")) &&
                        //swDocumentTypes_e Enumeration:swDocASSEMBLY--2;swDocPART--1;swDocDRAWING---3
                        //ModelDoc2 OpenDoc6( System.string FileName,System.int Type,System.int Options,
                        //                    System.string Configuration,out System.int Errors,out System.int Warnings)

                        if (test[test.Length - 1].ToUpper().Equals("SLDASM"))
                        {
                            fileType = 2;
                        }
                        else if (test[test.Length - 1].ToUpper().Equals("SLDPRT"))
                        {
                            fileType = 1;
                        }

                        SldWorks.ModelDoc2 swDoc = swApp.OpenDoc6(filePath,
                            fileType, 1, null, ref IErrors, ref IWarnings);



                        Debug.Print("FilePath: " + filePath);
                        Debug.Print("IErrors: " + IErrors);
                        Debug.Print("IWarnings: " + IWarnings);
                        //swDoc = swApp.ActiveDoc;
                        swDoc.Visible = true;


                        string valPartNumber = "";
                        string valoutPartNumber = "";
                        bool statusPartNumber;
                        string valCompNumber = "";
                        string valoutCompNumber = "";
                        bool statusCompNumber;
                        object[] configureNames;

                        ModelDocExtension swDocExtension = default(ModelDocExtension);
                        swDocExtension = swDoc.Extension;
                        CustomPropertyManager swDocProp = default(CustomPropertyManager);

                        configureNames = (object[])swDoc.GetConfigurationNames();
                        swDocProp = swDocExtension.get_CustomPropertyManager((string)configureNames[0]);

                        if (swDocProp != null)
                        {
                            statusPartNumber = swDocProp.Get4("图样代号", false, out valPartNumber, out valoutPartNumber);
                            statusCompNumber = swDocProp.Get4("部件代号", false, out valCompNumber, out valoutCompNumber);
                            Debug.Print("valoutPartNubmer: " + valoutPartNumber);
                            Debug.Print("valPartNubmer: " + valPartNumber);
                            Debug.Print("statusPartNumber:  " + statusPartNumber);
                            Debug.Print("valoutCompNumber:  " + valoutCompNumber);
                            Debug.Print("valCompNubmer: " + valCompNumber);
                            Debug.Print("statusCompNumber:  " + statusCompNumber);
                            if ((valCompNumber != "") && (valoutPartNumber != ""))
                            {
                                textBox1.AppendText(counter + ": " + filePath + "===<OK>" + System.Environment.NewLine);

                            }
                            else
                            {
                                counterError++;
                                if (valoutCompNumber == "")
                                {
                                    textBox1.AppendText(counter + ":WARNING（部件代号为空）: " + filePath + System.Environment.NewLine);
                                    textBox2.AppendText(counterError + " （部件代号为空）: " + filePath + System.Environment.NewLine);
                                }
                                if (valoutPartNumber == "")
                                {
                                    textBox1.AppendText(counter + ":WARNING（部件代号为空）: " + filePath + System.Environment.NewLine);
                                    textBox2.AppendText(counterError + " （图样代号为空）: " + filePath + System.Environment.NewLine);
                                }
                            }

                        }
                        else
                        {
                            textBox1.AppendText("Warning(swDocProp==null): " + (string)swDoc.GetTitle() + System.Environment.NewLine);
                            textBoxFileListPorp.AppendText("\"" + filePath + "\" ");
                        }

                        try
                        {
                            swApp.CloseDoc(swDoc.GetPathName());
                        }
                        catch (System.ComponentModel.Win32Exception we)
                        {

                            MessageBox.Show(this, we.Message);
                            return;
                            throw;
                        }
                    }
                    else
                    {
                        textBox2.AppendText("文件类型错误： " + filePath + System.Environment.NewLine);
                        textBoxFileListPorp.AppendText("-------------" + System.Environment.NewLine);

                    }

                }
                textBoxFileListPorp.AppendText("-------------" + System.Environment.NewLine);

                MessageBox.Show(this,"检查完成,共检查" + counter + "个文件");
                counter = 0;
                total = 0;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            searchPrinter();

        }

        #region checkDEBUG box
        //private void checkDEBUG_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (checkDEBUG.Checked == true)
        //    {
        //        buttonModifiyProp.Enabled = true;
        //    }
        //    else
        //    {
        //        buttonModifiyProp.Enabled = false;
        //    }
        //}
        #endregion

        private void buttonModifiyProp_Click(object sender, EventArgs e)
        {
            int sign = 0;

            if (dataGridView1.RowCount <= 1)
            {
                textBox2.AppendText("修改列表为空" + System.Environment.NewLine);
                sign = 1;
            }
            else
            {
                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                {
                    //if ((dataGridView1.Rows[i].Cells[0].Value.ToString() == "") ||
                    //(dataGridView1.Rows[i].Cells[0].Value == null))
                    if (dataGridView1.Rows[i].Cells[0].Value == null)
                    {
                        sign++;
                        textBox2.AppendText("第{" + (i + 1).ToString() + "}行存在空属性名称，请修正" + System.Environment.NewLine);
                        //MessageBox.Show(this,"第{" + (i + 1).ToString() + "}行存在空属性名称，请修正");
                    }
                }
            }

            if (sign != 0)
            {
                textBox2.AppendText("存在以上" + sign + "处错误，请修正后再试");
            }
            else
            {
                openFileDialog1.Filter = "SLDPRT, SLDASM files (*.SLDPRT,*.SLDASM)|*.sldprt;*.sldasm";
                openFileDialog1.Multiselect = true;
                counter = 0;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileNames = openFileDialog1.FileNames;
                    total = fileNames.Length;
                    countDown();
                    SldWorks.SldWorks swApp = new SldWorks.SldWorks();

                    foreach (string fileName in fileNames)
                    {
                        string[] test = fileName.Split('.');
                        filePath = fileName;
                        int fileType = 0;
                        counter++;
                        if ((test[test.Length - 1].ToUpper().Equals("SLDASM") || test[test.Length - 1].ToUpper().Equals("SLDPRT")) && fileName[0] != '~')
                        {

                            if (test[test.Length - 1].ToUpper().Equals("SLDASM"))
                            {
                                fileType = 2;
                            }
                            else if (test[test.Length - 1].ToUpper().Equals("SLDPRT"))
                            {
                                fileType = 1;
                            }

                            ModelDoc2 swDoc = swApp.OpenDoc6(filePath,
                                fileType, 1, null, ref IErrors, ref IWarnings);



                            Debug.Print("FilePath: " + filePath);
                            Debug.Print("IErrors: " + IErrors);
                            Debug.Print("IWarnings: " + IWarnings);
                            //swDoc = swApp.ActiveDoc;
                            swDoc.Visible = true;

                            object[] configureNames;

                            ModelDocExtension swDocExtension = default(ModelDocExtension);
                            swDocExtension = swDoc.Extension;
                            CustomPropertyManager swDocProp = default(CustomPropertyManager);

                            configureNames = (object[])swDoc.GetConfigurationNames();
                            swDocProp = swDocExtension.get_CustomPropertyManager((string)configureNames[0]);
                            //swDocProp = swDocExtension.get_CustomPropertyManager("");
                            string propName = "";
                            string propValue = "";
                            string porpOriginalValue = "";

                            if (swDocProp != null)
                            {
                                //dataGridView1.Rows[i].Cells[0].Value.ToString()
                                for (int i = 0; i < dataGridView1.RowCount - 1; i++)
                                {
                                    propName = dataGridView1.Rows[i].Cells[0].Value.ToString();
                                    //propValue = dataGridView1.Rows[i].Cells[1].Value.ToString();
                                    //propValue = (dataGridView1.Rows[i].Cells[1] == null ? "" : dataGridView1.Rows[i].Cells[1].Value.ToString());

                                    if (Convert.IsDBNull(dataGridView1.Rows[i].Cells[1]))
                                    {
                                        propValue = "";
                                    }
                                    else
                                    {
                                        propValue = dataGridView1.Rows[i].Cells[1].Value.ToString();
                                    }


                                    if (dataGridView1.Rows[i].Cells[2].Value != null)
                                    {
                                        if (dataGridView1.Rows[i].Cells[2] != null)
                                        {
                                            // find originalvalue and replace it with propValue
                                            string valValue = "";
                                            string valoutValue = "";
                                            bool status;
                                            porpOriginalValue = dataGridView1.Rows[i].Cells[2].Value.ToString();
                                            status = swDocProp.Get4(propName, false, out valValue, out valoutValue);
                                            if (valoutValue == porpOriginalValue)
                                            {
                                                swDocProp.Add3(propName, (int)swCustomInfoType_e.swCustomInfoText, propValue,
                                               (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                                                textBox1.AppendText((string)swDoc.GetTitle() + "属性{" + propName + "}由{" + valoutValue + "}替换为{" + propValue + "}" + System.Environment.NewLine);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        swDocProp.Add3(propName, (int)swCustomInfoType_e.swCustomInfoText, propValue,
                                            (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
                                        textBox1.AppendText((string)swDoc.GetTitle() + "增加属性{" + propName + "}值{" + propValue + "}" + System.Environment.NewLine);
                                    }
                                }

                            }
                            else
                            {
                                textBox1.AppendText("Warning(swDocProp==null): " + (string)swDoc.GetTitle() + System.Environment.NewLine);

                            }

                            try
                            {
                                swDoc.Save();
                                swApp.CloseDoc(swDoc.GetPathName());
                            }
                            catch (System.ComponentModel.Win32Exception we)
                            {

                                MessageBox.Show(this, we.Message);
                                return;
                                throw;
                            }
                        }
                        else
                        {
                            textBox2.AppendText("文件类型错误： " + filePath + System.Environment.NewLine);
                        }

                    }
                    MessageBox.Show(this,"编辑完成,共编辑" + counter + "个文件");
                    counter = 0;
                    total = 0;
                }
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(this,"确认清空属性编辑列表？", "提示", MessageBoxButtons.OKCancel);
            if (dr == DialogResult.OK)
            {
                dataGridView1.Rows.Clear();
            }

        }

        private void searchPrinter()
        {

            //defaultPrinter = print.PrinterSettings.PrinterName;

            foreach (String printerName in PrinterSettings.InstalledPrinters)
            {
                //    listPrinters.Add(printerName);
                if (printerName.IndexOf("PDF") != -1)
                {
                    printerPDF = printerName;

                }
                else if (printerName.IndexOf("1020") != -1)
                {
                    printerA4 = printerName;

                }
                else if (printerName.IndexOf("5200") != -1)
                {
                    printerA3 = printerName;

                }

                printerPDFComboBox.Items.Add(printerName);
                printerA3ComboBox.Items.Add(printerName);
                printerA4ComboBox.Items.Add(printerName);

            }
            if ((printerPDF != "") && (printerA3 != "") && (printerA4 != ""))
            {
                textBox3.AppendText("======================默认打印机========================" + System.Environment.NewLine);
                textBox3.AppendText("PDF打印机： " + printerPDF + System.Environment.NewLine);
                textBox3.AppendText("A4打印机： " + printerA4 + System.Environment.NewLine);
                textBox3.AppendText("A3打印机： " + printerA3 + System.Environment.NewLine);
                textBox3.AppendText("===============如有需要请重新选定打印机==================" + System.Environment.NewLine);
            }
            else
            {
                textBox4.AppendText("PDF打印机： " + printerPDF + System.Environment.NewLine);
                textBox4.AppendText("A4打印机： " + printerA4 + System.Environment.NewLine);
                textBox4.AppendText("A3打印机： " + printerA3 + System.Environment.NewLine);
            }
        }
        private void buttonPrint_Click(object sender, EventArgs e)
        {
            PrintSpecification printSpec;
            ModelDocExtension swModelDocExt;
            SldWorks.PageSetup swPageSetup;
            DrawingDoc swDrwng;
            string printerName = "";

            openFileDialog1.Filter = "SLDDRW files (*.slddrw)|*.slddrw";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNames = openFileDialog1.FileNames;
                total = fileNames.Length;
                countDown();
                SldWorks.SldWorks swApp = new SldWorks.SldWorks();

                foreach (string fileName in fileNames)
                {
                    string[] test = fileName.Split('.');

                    if (test[test.Length - 1].ToUpper().Equals("SLDDRW") && (fileName[0] != '~'))
                    {
                        filePath = fileName;

                        ModelDoc2 swDoc = swApp.OpenDoc6(filePath,
                            3, 1, null, ref IErrors, ref IWarnings);

                        Debug.WriteIf(swDoc == null, swDoc.ToString() + " IS NULL");
                        swModelDocExt = swDoc.Extension;

                        //////////////////////////////////////
                        #region GET Print Property
                        swDrwng = (DrawingDoc)swDoc;
                        Debug.WriteIf(swDrwng == null, swDoc.ToString() + "  " + swDrwng.ToString());
                        Sheet activeSheet = default(Sheet);
                        activeSheet = (Sheet)swDrwng.GetCurrentSheet();

                        sheetProperties = (double[])activeSheet.GetProperties();
                        paperSize = (swDwgPaperSizes_e)activeSheet.GetSize(ref width, ref height);

                        #endregion
                        Debug.Print(paperSize.ToString());

                        swPageSetup = (SldWorks.PageSetup)swDoc.PageSetup;
                        swPageSetup.ScaleToFit = true;
                        Debug.Print(swPageSetup.Orientation.ToString());
                        printSpec = (PrintSpecification)swModelDocExt.GetPrintSpecification();
                        printSpec.ConvertToHighQuality = true;
                        printSpec.AddPrintRange(1, 1);
                       // printSpec.FromScale = (double)sheetProperties[2];
                       // printSpec.ToScale = (double)sheetProperties[3];
                        printSpec.ScaleMethod = (int)SwConst.swPrintSelectionScaleFactor_e.swPrintSelection;
                        printSpec.PrinterQueue = "";
                        //// PRINT TO FILE
                        printSpec.PrintToFile = false;



                        Debug.Print("Printing specifications:");
                        Debug.Print("  Collate: " + printSpec.Collate);
                        Debug.Print("  Convert to high quality: " + printSpec.ConvertToHighQuality);
                        Debug.Print("  Current sheet: " + printSpec.CurrentSheet);
                        Debug.Print("  From scale: " + printSpec.FromScale);
                        Debug.Print("  To scale: " + printSpec.ToScale);
                        Debug.Print("  Number of copies: " + printSpec.NumberOfCopies);
                        Debug.Print("  Print background: " + printSpec.PrintBackground);
                        Debug.Print("  Print cross hatch on out-of-date views: " + printSpec.PrintCrossHatchOnOutOfDateViews);
                        Debug.Print("  Printer queue: " + printSpec.PrinterQueue);
                        Debug.Print("  Print white items black: " + printSpec.PrintWhiteItemsBlack);
                        Debug.Print("  Selection as defined in swPrintSelectionScaleFactor_e: " + printSpec.ScaleMethod);
                        Debug.Print("  Total sheet count: " + printSpec.SheetCount);
                        Debug.Print("  X origin: " + printSpec.XOrigin);
                        Debug.Print("  Y origin: " + printSpec.YOrigin);



                        //Print the drawing to your default printer
                        if (checkBoxPDF.Checked == true)
                        {
                            printerName = printerPDF;
                            if (paperSize == swDwgPaperSizes_e.swDwgPaperA4sizeVertical)
                            {
                                //swPageSetup.PrinterPaperSize = 9;   // A3 PaperSize = 8; A4 PaperSize = 9
                                //swPageSetup.Orientation = 1;

                                #region A4 MODIFY FOR SW2017
                                swPageSetup.PrinterPaperSize = 9;   // A3 PaperSize = 8; A4 PaperSize = 9
                                swPageSetup.Orientation = 1;
                                #endregion

                            }
                            else
                            {
                                //swPageSetup.PrinterPaperSize = 8;   // A3 PaperSize = 8; A4 PaperSize = 9
                                //swPageSetup.Orientation = 2;
                                #region A4 MODIFY FOR SW2017
                                swPageSetup.PrinterPaperSize = 8;   // A3 PaperSize = 8; A4 PaperSize = 9
                                swPageSetup.Orientation = 2;
                                #endregion

                            }
                            swModelDocExt.PrintOut4(printerPDF, "", printSpec);
                        }
                        else
                        {
                            if (paperSize == swDwgPaperSizes_e.swDwgPaperA3size || paperSize==swDwgPaperSizes_e.swDwgPaperA2size ||
                                paperSize== swDwgPaperSizes_e.swDwgPaperA1size || paperSize == swDwgPaperSizes_e.swDwgPaperA0size)
                            {
                                printerName = printerA3;
                                swPageSetup.PrinterPaperSize = 8;   // A3 PaperSize = 8; A4 PaperSize = 9 // MARK CHANGED FOR VERSION 2017
                                swPageSetup.Orientation = 2;    // 横向

                                swModelDocExt.PrintOut4(printerA3, "", printSpec);
                            }
                            else if (paperSize == swDwgPaperSizes_e.swDwgPaperA4sizeVertical)
                            {
                                printerName = printerA4;
                                swPageSetup.PrinterPaperSize = 9;   // A3 PaperSize = 8; A4 PaperSize = 9
                                swPageSetup.Orientation = 1;    // 纵向  // MARK CHANGED FOR VERSION 2017
                                //print.DefaultPageSettings.PaperSize = new PaperSize("Custom", 500, 300);

                                swModelDocExt.PrintOut4(printerA4, "", printSpec);

                            }
                        }

                        textBox3.AppendText((counter+1) + ": 正在" + printerName + "打印文档" + fileName+System.Environment.NewLine);
                        textBoxFileListPrint.AppendText("\"" + fileName + "\" ");


                        printSpec.RestoreDefaults();
                        printSpec.ResetPrintRange();


                        //choosePrinter(defaultPrinter);
                        swDoc.Visible = true;
                        counter++;
                        textBox1.AppendText(counter + ": " + filePath + System.Environment.NewLine);
                        try
                        {
                            swApp.CloseDoc(swDoc.GetPathName());
                        }
                        catch (System.ComponentModel.Win32Exception we)
                        {

                            MessageBox.Show(this, we.Message);
                            return;
                            throw;
                        }
                    }
                }
                textBoxFileListPrint.AppendText("-------------" + System.Environment.NewLine);
                MessageBox.Show(this,"打印完成,共打印" + counter + "个文件");

                counter = 0;
                total = 0;
            }



        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void buttonClose2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void printerSet_Click(object sender, EventArgs e)
        {
            if(printerPDFComboBox.SelectedItem == null)
            {
                MessageBox.Show(this,"请选定PDF打印机！！");
                textBox4.AppendText("请选定PDF打印机！！" + System.Environment.NewLine);

            }
            else if(printerA3ComboBox.SelectedItem==null)
            {
                MessageBox.Show(this,"请选定A3打印机！！");
                textBox4.AppendText("请选定A3打印机！！" + System.Environment.NewLine);
            }
            else if(printerA4ComboBox.SelectedItem==null)
            {
                MessageBox.Show(this,"请选定A4打印机！！");
                textBox4.AppendText("请选定A4打印机！！" + System.Environment.NewLine);
            }
            else
            {

                printerPDF = printerPDFComboBox.SelectedItem.ToString();
                printerA3 = printerA3ComboBox.SelectedItem.ToString();
                printerA4 = printerA4ComboBox.SelectedItem.ToString();
                textBox3.AppendText("=======================================================" + System.Environment.NewLine);
                textBox3.AppendText("PDF打印机： " + printerPDF + System.Environment.NewLine);
                textBox3.AppendText("A4打印机： " + printerA4 + System.Environment.NewLine);
                textBox3.AppendText("A3打印机： " + printerA3 + System.Environment.NewLine);
                textBox3.AppendText("=======================================================" + System.Environment.NewLine);
            }
        }

        private void buttonClearPrintList_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(this,"确认清空打印清单？", "提示", MessageBoxButtons.OKCancel);
            if (dr == DialogResult.OK)
            {
                textBox3.Text = "";
                textBoxFileListPrint.Text = "";
                textBox3.AppendText("======================默认打印机========================" + System.Environment.NewLine);
                textBox3.AppendText("PDF打印机： " + printerPDF + System.Environment.NewLine);
                textBox3.AppendText("A4打印机： " + printerA4 + System.Environment.NewLine);
                textBox3.AppendText("A3打印机： " + printerA3 + System.Environment.NewLine);
                textBox3.AppendText("===============如有需要请重新选定打印机==================" + System.Environment.NewLine);
                textBox4.Text = "";
                label1.Text = AUTHOR;
            }
        }

        private void buttonRenameRcdClear_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show(this,"确认清空文件更名清单？", "提示", MessageBoxButtons.OKCancel);
            if (dr == DialogResult.OK)
            {
                textBoxOrgnization.Text = "";
                textBoxOrgWarning.Text = "";
                label1.Text = AUTHOR;
            }
        }

        private void buttonRename_Click(object sender, EventArgs e)
        {
            if(radioButtonPart.Checked==true)
            {
                openFileDialog1.Filter = "零件SLDPRT files (*.SLDPRT)|*.sldprt";
                openFileDialog1.Multiselect = true;
            }
            else if(radioButtonAssm.Checked==true)
            {
                DialogResult dr = MessageBox.Show(this,"请先完成所有相关零件的重命名", "提示", MessageBoxButtons.OKCancel);
                if (dr == DialogResult.OK)
                {
                    openFileDialog1.Filter = "装配体SLDASM files (*.SLDASM)|*.sldasm";
                    openFileDialog1.Multiselect = true;
                }
            }
            else
            {
                DialogResult dr = MessageBox.Show(this,"请选择文件类型", "警告", MessageBoxButtons.OKCancel);
            }

            counter = 0;
            int counterError = 0;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNames = openFileDialog1.FileNames;
                total = fileNames.Length;
                countDown();
                SldWorks.SldWorks swApp = new SldWorks.SldWorks();
                int fileType = 0;

                foreach (string fileName in fileNames)
                {
                    string[] test = fileName.Split('.');
                    #region filePathTest
                    int s = test[0].LastIndexOf('\\');
                    string path = test[0].Substring(0, s + 1);
                    string extension = "";
                    #endregion
                    countDown();

                    filePath = fileName;

                    counter++;


                    #region GetNumber&Name
                    if ((test[test.Length - 1].ToUpper().Equals("SLDASM") || test[test.Length - 1].ToUpper().Equals("SLDPRT")) && fileName[0] != '~')
                    {
                        //(test[test.Length - 1].Equals("SLDASM") || test[test.Length - 1].Equals("SLDPRT")) &&
                        //swDocumentTypes_e Enumeration:swDocASSEMBLY--2;swDocPART--1;swDocDRAWING---3
                        //ModelDoc2 OpenDoc6( System.string FileName,System.int Type,System.int Options,
                        //                    System.string Configuration,out System.int Errors,out System.int Warnings)

                        if (test[test.Length - 1].ToUpper().Equals("SLDASM"))
                        {
                            fileType = 2;
                            extension = "." + "sldasm";
                        }
                        else if (test[test.Length - 1].ToUpper().Equals("SLDPRT"))
                        {
                            fileType = 1;
                            extension = "." + "sldprt";
                        }

                        ModelDoc2 swDoc = swApp.OpenDoc6(filePath,
                            fileType, 1, null, ref IErrors, ref IWarnings);



                        Debug.Print("Rename FilePath: " + filePath);
                        Debug.Print("Rename IErrors: " + IErrors);
                        Debug.Print("Rename IWarnings: " + IWarnings);
                        //swDoc = swApp.ActiveDoc;
                        swDoc.Visible = true;


                        string valPartNumber = "";
                        string valoutPartNumber = "";
                        bool statusPartNumber;
                        string valPartName = "";// valCompNumber = "";
                        string valoutPartName = "";// valoutCompNumber = "";
                        bool statusPartName;// statusCompNumber;
                        object[] configureNames;

                        ModelDocExtension swDocExtension = default(ModelDocExtension);
                        swDocExtension = swDoc.Extension;
                        CustomPropertyManager swDocProp = default(CustomPropertyManager);
                        string name = "";
                        configureNames = (object[])swDoc.GetConfigurationNames();
                        for (int i = 0; i < configureNames.Length; i++)
                        {
                            swDocProp = swDocExtension.get_CustomPropertyManager((string)configureNames[i]);

                            if (swDocProp != null)
                            {
                                statusPartNumber = swDocProp.Get4("图样代号", false, out valPartNumber, out valoutPartNumber);
                                statusPartName = swDocProp.Get4("图样名称", false, out valPartName, out valoutPartName);
                                Debug.Print("valoutPartNubmer: " + valoutPartNumber);
                                Debug.Print("valPartNubmer: " + valPartNumber);
                                Debug.Print("statusPartNumber:  " + statusPartNumber);
                                Debug.Print("valoutPartName:  " + valoutPartName);
                                Debug.Print("valCompNubmer: " + valPartName);
                                Debug.Print("statusPartName:  " + statusPartName);

                                if ((valPartName != "") && (valoutPartNumber != ""))
                                {
                                    name = valPartNumber + valPartName;
                                    textBoxOrgnization.AppendText(counter + ": " + name + "Changed  ===<OK>" + System.Environment.NewLine);
                                }
                                else
                                {
                                    counterError++;
                                    if (valoutPartName == "")
                                    {
                                        textBoxOrgnization.AppendText(counter + ":WARNING（部件代号为空）: " + filePath + System.Environment.NewLine);
                                        textBoxOrgWarning.AppendText(counterError + " （部件代号为空）: " + filePath + System.Environment.NewLine);
                                    }
                                    if (valoutPartNumber == "")
                                    {
                                        textBoxOrgnization.AppendText(counter + ":WARNING（部件代号为空）: " + filePath + System.Environment.NewLine);
                                        textBoxOrgWarning.AppendText(counterError + " （图样代号为空）: " + filePath + System.Environment.NewLine);
                                    }
                                }

                            }
                            else
                            {
                                textBoxOrgWarning.AppendText("Warning(swDocProp==null): " + (string)swDoc.GetTitle() + System.Environment.NewLine);

                            }
                        }
                        #endregion

                        try
                        {
                            string saveAsName = path + name + "." +extension;
                            if (name != "")
                            {
                                swDoc.SaveAs(saveAsName);
                                textBoxOrgnization.AppendText(saveAsName + System.Environment.NewLine);
                                if(fileType ==1)    // fileType==SLDPRT
                                {

                                    // find slddrw==true

                                    // reloadOrReplace Model of slddrw
                                    // Rename slddrw( saveAs)
                                }
                                else if(fileType == 2)  // fileType == SLDASM
                                {
                                    // Replace all models
                                    // save
                                    // fine slddrw == true
                                    // reloadOrReplace model of sldasm
                                    // rename slddrw(saveas)
                                }
                            }
                            else
                            {
                                textBoxOrgWarning.AppendText("!!SaveAs Error:: " + saveAsName + System.Environment.NewLine);
                            }

                            swApp.CloseDoc(swDoc.GetPathName());
                        }
                        catch (System.ComponentModel.Win32Exception we)
                        {

                            MessageBox.Show(this, we.Message);
                            return;
                            throw;
                        }
                    }
                    else
                    {
                        textBoxOrgWarning.AppendText("文件类型错误： " + filePath + System.Environment.NewLine);
                    }

                }
                MessageBox.Show(this,"检查完成,共检查" + counter + "个文件");
                counter = 0;
                total = 0;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //OPEN ASSOCIATED DRAWING TEST
            ModelDoc2 swModel;
            ModelDocExtension swModelExt;
            SldWorks.SldWorks swApp = new SldWorks.SldWorks();


            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModelExt = (ModelDocExtension)swModel.Extension;
            swModelExt.RunCommand((int)SwCommands.swCommands_e.swCommands_Open_Associated_Drw, "");
        }



        private void buttonSaveAsPDF_Click(object sender, EventArgs e)
        {
            ModelDocExtension swModelDocExt;
            DrawingDoc swDrwng;
            string pdfName = "";
            ExportPdfData swExportPDFData = default(ExportPdfData);
            bool boolstatus = false;
            int errors = 0;
            int warnings = 0;



            openFileDialog1.Filter = "SLDDRW files (*.slddrw)|*.slddrw";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNames = openFileDialog1.FileNames;
                total = fileNames.Length;
                countDown();
                SldWorks.SldWorks swApp = new SldWorks.SldWorks();


                foreach (string fileName in fileNames)
                {
                    // string[] test = fileName.Split('.');

                    int clipPoint = fileName.LastIndexOf('.');
                    string pathName = fileName.Substring(0, fileName.LastIndexOf('.'));
                    string extensionName = fileName.Remove(0, fileName.Length - 6);

                    string[] test = { pathName, extensionName };


                    if (test[test.Length - 1].ToUpper().Equals("SLDDRW") && (fileName[0] != '~'))
                    {
                        string valPartNumber = "";
                        string valoutPartNumber = "";
                        string valPartName = "";
                        string valoutPartName = "";
                        filePath = fileName;
                        pdfName = test[0] + ".PDF";

                        ModelDoc2 swDoc = swApp.OpenDoc6(filePath,
                            3, 1, null, ref IErrors, ref IWarnings);

                        Debug.WriteIf(swDoc == null, swDoc.ToString() + " IS NULL");
                        swModelDocExt = swDoc.Extension;


                        swDrwng = (DrawingDoc)swDoc;
                        Debug.WriteIf(swDrwng == null, swDoc.ToString() + "  " + swDrwng.ToString());
                        Sheet activeSheet = default(Sheet);
                        activeSheet = (Sheet)swDrwng.GetCurrentSheet();
                        if (checkBoxUnmodifyPDFName.Checked == false)
                        {
                            #region SetPDFName
                            object[] swPart = swDoc.GetDependencies2(true, true, true);
                            foreach (var o in swPart)
                            {
                                Debug.Print((string)o);
                            }

                            //////////////////////////需要判断是PART还是ASSEMBLY///////////
                            ModelDoc2 swDocDepend = swApp.GetOpenDocumentByName((string)swPart[1]);
                            if (swDocDepend != null)
                            {

                                Debug.Print("swPart[1]: " + (string)swPart[1]);
                                Debug.Print("IErrors: " + IErrors);
                                Debug.Print("IWarnings: " + IWarnings);


                                #region Get Properties

                                ModelDocExtension swDocExtension = default(ModelDocExtension);
                                swDocExtension = swDocDepend.Extension;
                                CustomPropertyManager swDocProp = default(CustomPropertyManager);

                                var configureNames = (object[])swDocDepend.GetConfigurationNames();
                                swDocProp = swDocExtension.CustomPropertyManager[(string)configureNames[0]];

                                if (swDocProp != null)
                                {
                                    var statusPartNumber = swDocProp.Get4("图样代号", false, out valPartNumber, out valoutPartNumber);
                                    var statusPartName = swDocProp.Get4("图样名称", false, out valPartName, out valoutPartName);

                                    Debug.Print("valoutPartNubmer: " + valoutPartNumber);
                                    Debug.Print("valPartNubmer: " + valPartNumber);
                                    Debug.Print("statusPartNumber:  " + statusPartNumber);

                                    if (valPartName == "")
                                    {
                                        textBox3.AppendText("WARNING（图样名称为空）: " + filePath + System.Environment.NewLine);
                                        textBox4.AppendText("WARNING（图样名称为空）: " + filePath + System.Environment.NewLine);
                                    }
                                    if (valoutPartNumber == "")
                                    {
                                        textBox3.AppendText("WARNING（图样代号为空）: " + filePath + System.Environment.NewLine);
                                        textBox4.AppendText("WARNING（图样代号为空）: " + filePath + System.Environment.NewLine);
                                    }
                                    #endregion
                                }
                                else
                                {
                                    textBox4.AppendText("Warning(swDocProp==null): " + (string)swDoc.GetTitle());
                                }
                            }
                            else
                            {
                                textBox4.AppendText("找不到参考文件" + (string)swPart[1]);
                            }
                            #endregion
                        }else
                        {
                            valPartName = "";
                            valPartNumber = "";
                        }

                        if (valPartName != "" & valPartNumber != "")
                        {
                            // MODIFY PDF NAME
                            int s = test[0].LastIndexOf('\\');
                            string path = test[0].Substring(0, s + 1);
                            string name = valPartNumber + valPartName;

                            // create pdf dir
                            string pdfDirPath = path + "pdf\\";
                            if (!Directory.Exists(pdfDirPath))
                            {
                                DirectoryInfo directoryInfo = new DirectoryInfo(pdfDirPath);
                                directoryInfo.Create();

                            }

                            pdfName = pdfDirPath + name + "." + "pdf";
                        }
                        else  // PartName & PartNumber有一为空时，PDF文件以原SLDDRW文件名命名
                        {
                            int s = test[0].LastIndexOf('\\');
                            string path = test[0].Substring(0, s + 1);
                            //string[] test = fileName.Split('.');
                            string[] temp = test[0].Split('\\');
                            string name = temp[temp.Length - 1];
                            string pdfDirPath = path + "pdf\\";
                            if (!Directory.Exists(pdfDirPath))
                            {
                                DirectoryInfo directoryInfo = new DirectoryInfo(pdfDirPath);
                                directoryInfo.Create();

                            }

                            pdfName = pdfDirPath + name + "." + "pdf";
                        }

                        sheetProperties = (double[])activeSheet.GetProperties();
                        swExportPDFData = (ExportPdfData)swApp.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);

                        swExportPDFData.ViewPdfAfterSaving = false;
                        boolstatus = swModelDocExt.SaveAs(pdfName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, swExportPDFData, ref errors, ref warnings);

                        textBox3.AppendText((counter + 1) + ": 正在将 " + fileName + "转换为PDF文档" + System.Environment.NewLine);
                        textBoxFileListPrint.AppendText("\""+fileName+ "\" ");


                        //choosePrinter(defaultPrinter);
                        swDoc.Visible = true;
                        counter++;
                        textBox1.AppendText(counter + ": " + filePath + System.Environment.NewLine);
                        try
                        {
                            swApp.CloseDoc(swDoc.GetPathName());
                        }
                        catch (System.ComponentModel.Win32Exception we)
                        {

                            MessageBox.Show(this, we.Message);
                            return;
                            throw;
                        }
                    }
                }
                textBoxFileListPrint.AppendText("-------------" + System.Environment.NewLine);
                MessageBox.Show(this,"转换PDF完成,共生成" + counter + "个PDF文件");

                counter = 0;
                total = 0;
            }
        }


        //将模型另存到STEP中性格式
        private void buttonSaveAs_Click(object sender, EventArgs e)
        {
            ModelDocExtension swModelDocExt;
            //DrawingDoc swDrwng;
            string stepFileName = "";
            //ExportPdfData swExportPDFData = default(ExportPdfData);
            bool boolstatus = false;
            int errors = 0;
            int warnings = 0;



            openFileDialog1.Filter = "SLDPRT, SLDASM files (*.SLDPRT,*.SLDASM)|*.sldprt;*.sldasm";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileNames = openFileDialog1.FileNames;
                total = fileNames.Length;
                countDown();
                SldWorks.SldWorks swApp = new SldWorks.SldWorks();


                foreach (string fileName in fileNames)
                {
                    // string[] test = fileName.Split('.');

                    int clipPoint = fileName.LastIndexOf('.');
                    string pathName = fileName.Substring(0, fileName.LastIndexOf('.'));
                    string extensionName = fileName.Remove(0, fileName.Length - 6);

                    string[] test = { pathName, extensionName };


                    if (((test[test.Length - 1].ToUpper().Equals("SLDPRT"))||
                        test[test.Length - 1].ToUpper().Equals("SLDASM") )&& (fileName[0] != '~'))
                    {
                        string valPartNumber = "";
                        string valoutPartNumber = "";
                        string valPartName = "";
                        string valoutPartName = "";
                        filePath = fileName;
                        stepFileName = test[0] + ".step";
                        int fileType = 0;

                        if (test[test.Length - 1].ToUpper().Equals("SLDASM"))
                        {
                            fileType = 2;
                        }
                        else if (test[test.Length - 1].ToUpper().Equals("SLDPRT"))
                        {
                            fileType = 1;
                        }

                        ModelDoc2 swDoc = swApp.OpenDoc6(filePath,
                             fileType, 1, null, ref IErrors, ref IWarnings);

                        Debug.WriteIf(swDoc == null, swDoc.ToString() + " IS NULL");
                        swModelDocExt = swDoc.Extension;


                        //swDrwng = (DrawingDoc)swDoc;
                        //Debug.WriteIf(swDrwng == null, swDoc.ToString() + "  " + swDrwng.ToString());
                        //Sheet activeSheet = default(Sheet);
                        //activeSheet = (Sheet)swDrwng.GetCurrentSheet();

                        #region SetStpFileName
                        //object[] swPart = swDoc.GetDependencies2(true, true, true);
                        //foreach (var o in swPart)
                        //{
                        //    Debug.Print((string)o);
                        //}

                        //////////////////////////需要判断是PART还是ASSEMBLY///////////
                        //ModelDoc2 swDocDepend = swApp.GetOpenDocumentByName((string)swPart[1]);
                        ModelDoc2 swDocDepend = swDoc;
                        if (swDocDepend != null)
                        {

                            //Debug.Print("swPart[1]: " + (string)swPart[1]);
                            Debug.Print("IErrors: " + IErrors);
                            Debug.Print("IWarnings: " + IWarnings);


                            #region Get Properties

                            ModelDocExtension swDocExtension = default(ModelDocExtension);
                            swDocExtension = swDocDepend.Extension;
                            CustomPropertyManager swDocProp = default(CustomPropertyManager);

                            var configureNames = (object[])swDocDepend.GetConfigurationNames();
                            swDocProp = swDocExtension.CustomPropertyManager[(string)configureNames[0]];

                            if (swDocProp != null)
                            {
                                var statusPartNumber = swDocProp.Get4("图样代号", false, out valPartNumber, out valoutPartNumber);
                                var statusPartName = swDocProp.Get4("图样名称", false, out valPartName, out valoutPartName);

                                Debug.Print("valoutPartNubmer: " + valoutPartNumber);
                                Debug.Print("valPartNubmer: " + valPartNumber);
                                Debug.Print("statusPartNumber:  " + statusPartNumber);

                                if (valPartName == "")
                                {
                                    textBox3.AppendText("WARNING（图样名称为空）: " + filePath + System.Environment.NewLine);
                                    textBox4.AppendText("WARNING（图样名称为空）: " + filePath + System.Environment.NewLine);
                                }
                                if (valoutPartNumber == "")
                                {
                                    textBox3.AppendText("WARNING（图样代号为空）: " + filePath + System.Environment.NewLine);
                                    textBox4.AppendText("WARNING（图样代号为空）: " + filePath + System.Environment.NewLine);
                                }
                                #endregion
                            }
                            else
                            {
                                textBox4.AppendText("Warning(swDocProp==null): " + (string)swDoc.GetTitle());
                            }
                        }
                        else
                        {
                           // textBox4.AppendText("找不到参考文件" + (string)swPart[1]);
                        }
                        #endregion

                        if (valPartName != "" & valPartNumber != "")
                        {
                            // MODIFY PDF NAME
                            int s = test[0].LastIndexOf('\\');
                            string path = test[0].Substring(0, s + 1);
                            string name = valPartNumber + valPartName;

                            // create pdf dir
                            string stepFileDirPath = path + "step\\";
                            if (!Directory.Exists(stepFileDirPath))
                            {
                                DirectoryInfo directoryInfo = new DirectoryInfo(stepFileDirPath);
                                directoryInfo.Create();

                            }

                            stepFileName = stepFileDirPath + name + "." + "step";
                        }
                        else  // PartName & PartNumber有一为空时，文件以原SLDDRW文件名命名
                        {
                            int s = test[0].LastIndexOf('\\');
                            string path = test[0].Substring(0, s + 1);
                            //string[] test = fileName.Split('.');
                            string[] temp = test[0].Split('\\');
                            string name = temp[temp.Length - 1];
                            string stepFileDirPath = path + "step\\";
                            if (!Directory.Exists(stepFileDirPath))
                            {
                                DirectoryInfo directoryInfo = new DirectoryInfo(stepFileDirPath);
                                directoryInfo.Create();

                            }

                            stepFileName = stepFileDirPath + name + "." + "step";
                        }


                        //boolstatus = swModelDocExt.SaveAs(stepFileName, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);
                        boolstatus = swModelDocExt.SaveAs(stepFileName, 0, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref errors, ref warnings);

                        textBox3.AppendText((counter + 1) + ": 正在将 " + fileName + "转换为STEP文档" + System.Environment.NewLine);
                        textBoxFileListPrint.AppendText("\"" + fileName + "\" ");



                        //choosePrinter(defaultPrinter);
                        swDoc.Visible = true;
                        counter++;
                        textBox1.AppendText(counter + ": " + filePath + System.Environment.NewLine);
                        try
                        {
                            swApp.CloseDoc(swDoc.GetPathName());
                        }
                        catch (System.ComponentModel.Win32Exception we)
                        {

                            MessageBox.Show(this, we.Message);
                            return;
                            throw;
                        }
                    }
                }
                textBoxFileListPrint.AppendText("-------------" + System.Environment.NewLine);
                MessageBox.Show(this, "转换STEP完成,共生成" + counter + "个STEP文件");

                counter = 0;
                total = 0;
            }

        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        #region CreateTable
        private void buttonCreateTable_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("请先打开目标装配体", "提示", MessageBoxButtons.OKCancel);
            if(dr == DialogResult.OK)
            {
                row = 2;
                sheetName = "BOM";
                int column = 1;

                //创建excel表格
                excel = new ExcelEdit();
                excel.Create();
                excel.AddSheet(sheetName);
                excel.SetCellValue(sheetName, 1, column++, "行号");
                excel.SetCellValue(sheetName, 1, column++, "序号");
                excel.SetCellValue(sheetName, 1, column++, "层级");
                excel.SetCellValue(sheetName, 1, column++, "图样代号");
                excel.SetCellValue(sheetName, 1, column++, "图样名称");
                excel.SetCellValue(sheetName, 1, column++, "目标名称");
                excel.SetCellValue(sheetName, 1, column++, "原始名称");
                excel.SetCellValue(sheetName, 1, column++, "目标路径");

                //递归读取装配体内容
                TraverseMain();

                //excel表格调整
                excel.wb.ActiveSheet.Columns["A:Z"].EntireColumn.AutoFit();
                int clipPoint = xlsName.LastIndexOf('.');
                xlsName = xlsName.Substring(0, xlsName.LastIndexOf('.')) + "BOM.xls";
                excel.SaveAs(xlsName);


                Debug.Print("main()over");

            }
            else if(dr == DialogResult.Cancel)
            {

            }
        }

        #region ForButtonOrgnization
        #region TraverseFeatureFeatures
        public void TraverseFeatureFeatures(Feature swFeat, long nLevel)
        {
            Feature swSubFeat;
            Feature swSubSubFeat;
            Feature swSubSubSubFeat;
            string sPadStr = " ";
            long i = 0;

            for (i = 0; i <= nLevel; i++)
            {
                sPadStr = sPadStr + " ";
            }
            while ((swFeat != null))
            {
                Debug.Print(sPadStr + swFeat.Name + " [" + swFeat.GetTypeName2() + "]");
                textBoxOrgnization.AppendText("<" + i + ">" + sPadStr + swFeat.Name + " [" + swFeat.GetTypeName2() + "]\r\n");
                swSubFeat = (Feature)swFeat.GetFirstSubFeature();

                while ((swSubFeat != null))
                {
                    Debug.Print(sPadStr + "  " + swSubFeat.Name + " [" + swSubFeat.GetTypeName() + "]");
                    textBoxOrgnization.AppendText("<" + i + ">" + sPadStr + "  " + swSubFeat.Name + " [" + swSubFeat.GetTypeName() + "]\r\n");
                    swSubSubFeat = (Feature)swSubFeat.GetFirstSubFeature();

                    while ((swSubSubFeat != null))
                    {
                        Debug.Print(sPadStr + "    " + swSubSubFeat.Name + " [" + swSubSubFeat.GetTypeName() + "]");
                        textBoxOrgnization.AppendText("<" + i + ">" + sPadStr + "    " + swSubSubFeat.Name + " [" + swSubSubFeat.GetTypeName() + "]\r\n");
                        swSubSubSubFeat = (Feature)swSubSubFeat.GetFirstSubFeature();

                        while ((swSubSubSubFeat != null))
                        {
                            Debug.Print(sPadStr + "      " + swSubSubSubFeat.Name + " [" + swSubSubSubFeat.GetTypeName() + "]");
                            textBoxOrgnization.AppendText("<" + i + ">" + sPadStr + "      " + swSubSubSubFeat.Name + " [" + swSubSubSubFeat.GetTypeName() + "]\r\n");
                            swSubSubSubFeat = (Feature)swSubSubSubFeat.GetNextSubFeature();

                        }

                        swSubSubFeat = (Feature)swSubSubFeat.GetNextSubFeature();

                    }

                    swSubFeat = (Feature)swSubFeat.GetNextSubFeature();

                }

                swFeat = (Feature)swFeat.GetNextFeature();

            }

        }
        #endregion

        #region TraverseComponentFeatures
        public void TraverseComponentFeatures(Component2 swComp, long nLevel)
        {
            Feature swFeat;

            swFeat = (Feature)swComp.FirstFeature();
            TraverseFeatureFeatures(swFeat, nLevel);
        }
        #endregion

        #region TraverseComponent
        //遍历装配体，将属性写入excel表格
        public void TraverseComponent(Component2 swComp, long nLevel)
        {
            object[] vChildComp;
            Component2 swChildComp;
            string sPadStr = " ";
            long i = 0;
            ModelDoc2 componentModelDoc2;
            int componentType;
            string originalPath;
            string targetPath;
            int columnNum = 0;
            valPartNumber = "";
            valoutPartNumber = "";
            valPartName = "";
            valoutPartName = "";
            string targetName = "";
            string drawingPath = "";//工程图路径
            string drawingName = "";//工程图名称，需带上扩展名

            ModelDocExtension componentExtension = default(ModelDocExtension);
            CustomPropertyManager componentCustomPropertyManager = default(CustomPropertyManager);

            vChildComp = (object[])swComp.GetChildren();
            for (i = 0; i < vChildComp.Length; i++)
            {
                swChildComp = (Component2)vChildComp[i];
                //test whether a part is a Toolbox part.
                //part = (ModelDoc2)swApp.ActiveDoc;
                //modelDocExt = part.Extension;
                //ret = modelDocExt.ToolboxPartType;
                //swNotAToolboxPart 0x0 = Not a Toolbox part
                //swToolboxCopiedPart 0x2 = Copied part
                //swToolboxStandardPart 0x1 = Standard part

                if (swChildComp.IsSuppressed() == false)//排除已压缩的特征
                {
                    componentModelDoc2 = swChildComp.GetModelDoc2();
                    if (componentModelDoc2 == null)
                    {
                        if (swChildComp == null)
                        {
                            Debug.Print("swChildComp == null>>" + vChildComp[i].ToString());
                        }
                        else
                        {
                            Debug.Print("componentModelDoc2 == null>>" + swChildComp.GetPathName());
                        }
                    }
                    Debug.Print(componentModelDoc2.GetPathName());
                    componentExtension = componentModelDoc2.Extension;
                    configureNames = (object[])componentModelDoc2.GetConfigurationNames();

                    componentType = componentExtension.ToolboxPartType;//获取组件类型（是否为Toolbox库中标准件）
                    originalPath = componentModelDoc2.GetPathName();
                    textBoxOrgnization.AppendText(componentType.ToString() + "::" + originalPath + "\r\n");

                    if (componentType == (int)(swToolBoxPartType_e.swNotAToolboxPart))//（排除Toolbox零件）
                    {
                        Debug.Print(sPadStr + "+" + swChildComp.Name2 + " <" + swChildComp.ReferencedConfiguration + ">");
                        textBoxOrgnization.AppendText("<<" + i + ":" + nLevel + ">>" + /*sPadStr + "+" +*/ swChildComp.Name2 + " <" + swChildComp.ReferencedConfiguration + ">\r\n");

                        //swChildComp.get
                        //Get properties
                        componentCustomPropertyManager = componentExtension.get_CustomPropertyManager((string)configureNames[0]);
                        if (componentCustomPropertyManager != null)
                        {
                            statusPartNumber = componentCustomPropertyManager.Get4("图样代号", false, out valPartNumber, out valoutPartNumber);
                            statusPartName = componentCustomPropertyManager.Get4("图样名称", false, out valPartName, out valoutPartName);
                            Debug.Print("valoutPartNubmer: " + valoutPartNumber);
                            Debug.Print("valPartNubmer: " + valPartNumber);
                            Debug.Print("statusPartNumber:  " + statusPartNumber);
                            Debug.Print("valoutPartName:  " + valoutPartName);
                            Debug.Print("valCompNubmer: " + valPartName);
                            Debug.Print("statusPartName:  " + statusPartName);

                            //创建目标文件路径
                            //对于有图样代号及图样名称的零件，文件名改为图样代号+图样名称形式，扩展名不变
                            if (valPartName != "" & valPartNumber != "")
                            {
                                int clipPoint = originalPath.LastIndexOf('.');
                                string pathName = originalPath.Substring(0, originalPath.LastIndexOf('.'));
                                string extensionName = originalPath.Remove(0, originalPath.Length - 6);
                                int s = pathName.LastIndexOf('\\');
                                path = pathName.Substring(0, s + 1);

                                targetName = valPartNumber + valPartName + "." + extensionName;
                                targetPath = path + targetName;
                            }
                            else
                            {
                                //对于无图样代号或无图样名称的零件，文件名不变
                                targetPath = originalPath;
                            }
                        }
                        else
                        {
                            //对于无图样代号或无图样名称的零件，文件名不变
                            targetPath = originalPath;
                        }
                        textBoxOrgnization.AppendText(valPartNumber + " " + valPartName);
                        textBoxOrgnization.AppendText("||" + targetPath + "\r\n");

                        #region ExcelWriting
                        //ExcelWriting
                        columnNum = 1;//每行起始写入的列号
                        excel.SetCellValue(sheetName, row, columnNum++, row-1);   //写入行号
                        excel.SetCellValue(sheetName, row, columnNum++, i);
                        excel.SetCellValue(sheetName, row, columnNum++, nLevel);
                        excel.SetCellValue(sheetName, row, columnNum++, valoutPartNumber);
                        excel.SetCellValue(sheetName, row, columnNum++, valoutPartName);
                        excel.SetCellValue(sheetName, row, columnNum++, targetName);
                        excel.SetCellValue(sheetName, row, columnNum++, originalPath);
                        excel.SetCellValue(sheetName, row, columnNum, targetPath);
                        row++;
                        #endregion
                        //TraverseComponentFeatures(swChildComp, nLevel);
                        TraverseComponent(swChildComp, nLevel + 1);
                    }
                    else
                    {
                        //Component不是普通零件，即忽略标准件部分
                        textBoxOrgnization.AppendText("[" + i + ":" + nLevel + "]" + /*sPadStr + "+" +*/ swChildComp.Name2 + " <" + swChildComp.ReferencedConfiguration + ">\r\n");
                    }
                }
            }
        }
        #endregion

        #region TraverseModelFeatures
        public void TraverseModelFeatures(ModelDoc2 swModel, long nLevel)
        {
            Feature swFeat;

            swFeat = (Feature)swModel.FirstFeature();
            TraverseFeatureFeatures(swFeat, nLevel);
        }
        #endregion

        public void TraverseMain()
        {

            ModelDoc2 swModel;
            ConfigurationManager swConfMgr;
            Configuration swConf;
            Component2 swRootComp;
            SldWorks.SldWorks swApp = new SldWorks.SldWorks();

            swModel = (ModelDoc2)swApp.ActiveDoc;
            swConfMgr = (ConfigurationManager)swModel.ConfigurationManager;
            swConf = (Configuration)swConfMgr.ActiveConfiguration;
            swRootComp = (Component2)swConf.GetRootComponent();

            System.Diagnostics.Stopwatch myStopwatch = new Stopwatch();
            myStopwatch.Start();




            xlsName = swModel.GetPathName();
            Debug.Print("File = " + xlsName);

            //TraverseModelFeatures(swModel, 1);

            if (swModel.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                TraverseComponent(swRootComp, 1);
            }

            myStopwatch.Stop();
            TimeSpan myTimespan = myStopwatch.Elapsed;
            Debug.Print("Time = " + myTimespan.TotalSeconds + " sec");

        }

        #endregion
        #endregion
        //根据表格整理文件名称
        private void buttonOrgnization_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("请先生成表格", "提示", MessageBoxButtons.OKCancel);
            if (dr == DialogResult.OK)
            {
                //生成表格文件路径
                ModelDoc2 swModel;
                SldWorks.SldWorks swApp = new SldWorks.SldWorks();
                swModel = (ModelDoc2)swApp.ActiveDoc;
                xlsName = swModel.GetPathName();
                int clipPoint = xlsName.LastIndexOf('.');
                xlsName = xlsName.Substring(0, xlsName.LastIndexOf('.')) + "BOM.xls";//默认表格文件名称

                //用于Pack&Go
                ConfigurationManager swConfMgr;
                Configuration swConf;
                Component2 swRootComp;
                ModelDocExtension swModelDocExt = default(ModelDocExtension);
                PackAndGo swPackAndGo = default(PackAndGo);
                swModel = (ModelDoc2)swApp.ActiveDoc;
                swConfMgr = (ConfigurationManager)swModel.ConfigurationManager;
                swConf = (Configuration)swConfMgr.ActiveConfiguration;
                swRootComp = (Component2)swConf.GetRootComponent();
                swModelDocExt = (ModelDocExtension)swModel.Extension;
                bool status = false;
                int i = 0;
                int namesCount = 0;

                //计时器
                System.Diagnostics.Stopwatch myStopwatch = new Stopwatch();
                myStopwatch.Start();

                swPackAndGo = (PackAndGo)swModelDocExt.GetPackAndGo();



                //判断表格是否存在
                if (File.Exists(xlsName))
                {
                    //ExcelEdit excelRead = new ExcelEdit();
                    //excelRead.Open(xlsName);
                    string originalPath;
                    string targetPath;
                    int row=2;
                    int lineNumber = 0;
                    int column = 1;
                    int serialNumber = 0;
                    int level = 0;
                    string partName;
                    string partNumber;
                    string path;
                    string targetName;
                    string targetExtensionname;

                    Microsoft.Office.Interop.Excel._Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = true;
                    ExcelEdit ee = new ExcelEdit();
                    ee.wb = excelApp.Workbooks.Open(xlsName);//打开workbook
                    ee.sheetName = "BOM";//设定worksheet名称
                    ee.ws = (Worksheet)ee.wb.Worksheets[ee.sheetName];
                    string l = ee.Read(1, 2);
                    Debug.Print(l+"ee.Read");

                    // Get number of documents in assembly
                    namesCount = swPackAndGo.GetDocumentNamesCount();
                    Debug.Print("  Number of model documents: " + namesCount);
                    // Include any drawings, SOLIDWORKS Simulation results, and SOLIDWORKS Toolbox components
                    swPackAndGo.IncludeDrawings = true;
                    Debug.Print(" Include drawings: " + swPackAndGo.IncludeDrawings);
                    swPackAndGo.IncludeSimulationResults = true;
                    Debug.Print(" Include SOLIDWORKS Simulation results: " + swPackAndGo.IncludeSimulationResults);
                    swPackAndGo.IncludeToolboxComponents = false;
                    Debug.Print(" Include SOLIDWORKS Toolbox components: " + swPackAndGo.IncludeToolboxComponents);
                    object fileNames;
                    object[] pgFileNames = new object[namesCount - 1];
                    status = swPackAndGo.GetDocumentNames(out fileNames);
                    pgFileNames = (object[])fileNames;

                    Debug.Print("");
                    Debug.Print("  Current path and filenames: ");
                    if ((pgFileNames != null))
                    {
                        for (i = 0; i <= pgFileNames.GetUpperBound(0); i++)
                        {
                            Debug.Print("    The path and filename is: " + pgFileNames[i]);
                        }
                    }

                    // Flatten the Pack and Go folder structure; save all files to the root directory
                    swPackAndGo.FlattenToSingleFolder = true;

                    // Verify document paths and filenames after adding prefix and suffix
                    object getFileNames;
                    object getDocumentStatus;
                    string[] pgGetFileNames = new string[namesCount - 1];
                    status = swPackAndGo.GetDocumentSaveToNames(out getFileNames, out getDocumentStatus);
                    pgGetFileNames = (string[])getFileNames;
                    Debug.Print("");
                    Debug.Print("名称变更后：");

                    //此for循环用于遍历及变更文件名称
                    for (i = 0; i <= namesCount - 1; i++)
                    {
                        //rename
                        //数据读取至出现空行号
                        while ((ee.Read(row, column) != null) && (ee.Read(row, column) != ""))
                        {
                            lineNumber = Convert.ToInt32(ee.Read(row, column++));
                            serialNumber = Convert.ToInt32(ee.Read(row, column++));
                            level = Convert.ToInt32(ee.Read(row, column++));
                            partNumber = ee.Read(row, column++);
                            partName = ee.Read(row, column++);
                            targetName = ee.Read(row, column++);
                            originalPath = ee.Read(row, column++);
                            targetPath = ee.Read(row, column);

                            Debug.Print(row + ":" + column + "::" + lineNumber + ":" + serialNumber + ":" +
                                level + ":" + partNumber + ":" + partName + ":" + originalPath + ":" +
                                targetPath);
                            column = 1;//列号回滚至行首


                            #region 重命名操作
                            //IPackAndGo??
                            //打开文件
                            //重命名==检查名称不许变更的不做处理
                            //更新属性==表格增加属性变更列，已填写的可以写入新属性？
                            //更新参考 ======同步更新图纸参考======更新图纸名称
                            //图纸slddrw名称尚未能写入excel
                            //注意备份
                            if (((pgGetFileNames[i]).ToUpper() == originalPath.ToUpper())&&(targetName!="")&&(targetName!=null))
                            {
                                Debug.Print("pgGetFileNames["+i+"] = " + pgGetFileNames[i]);

                                pgGetFileNames[i] = targetName;

                            }
                            #endregion
                            row++;//下移一行
                        }

                        Debug.Print("    My path and filename is: " + pgGetFileNames[i]);
                    }


                    //设置另存为的文件名
                    //********字符串中似乎仅截取文件名部分，路径部分自动替换为SetSaveToName所指定的路径？***
                    swPackAndGo.SetDocumentSaveToNames(pgGetFileNames);
                    //设置另存为的路径
                    //status = swPackAndGo.SetSaveToName(true, myPath);
                    // Pack and Go
                    int[] statuses = (int[])swModelDocExt.SavePackAndGo(swPackAndGo);

                    //计时器停止，输出时间
                    myStopwatch.Stop();
                    TimeSpan myTimespan = myStopwatch.Elapsed;
                    Debug.Print("Time = " + myTimespan.TotalSeconds + " sec");

                }
                else
                {
                    //未找到表格
                    MessageBox.Show("未找到"+ xlsName, "警告", MessageBoxButtons.OK);
                }


            }
            else
            {
                //未生成表格时不做处理
            }

        }
    }
}
