using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 刷新文件卡
{
    public partial class FormAddProperty : Form
    {
        public FormAddProperty()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void buttonReturn_Click(object sender, EventArgs e)
        {
            string dataGridValues = dataGridView1.DataMember;
            Close();
            
        }

        private void FormAddProperty_Load(object sender, EventArgs e)
        {

        }

        private int counter = 0;
        private string fileNames [];
        private void buttonAddProperty_Click(object sender, EventArgs e)
        {
            openFileDlgAddPro.Filter = "SLDPRT, SLDASM files (*.SLDPRT,*.SLDASM)|*.sldprt;*.sldasm";
            openFileDlgAddPro.Multiselect = true;
            counter = 0;
            int counterError = 0;
            if (openFileDlgAddPro.ShowDialog() == DialogResult.OK)
            {
                fileNames = openFileDlgAddPro.FileNames;
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

                        ModelDoc2 swDoc = swApp.OpenDoc6(filePath,
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
                    }

                }
                MessageBox.Show("检查完成,共检查" + counter + "个文件");
                counter = 0;
                total = 0;
            }
        }

    }
}
}
