using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geoprocessor;
using System.Data.OleDb;
using System.IO;
using System.Xml;
using System.Reflection;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading; 


namespace ExcelToShape
{
    public partial class FormMain : Form
    {
        //string[] fields = {"XH","XMMC","XMSZD","JSQ","ZGM","ZTZ","XZGDMJ","SJKYYZBP","SHSJ",
        //                         "YSRQ","JSXMMC","GGMJ","JSYDZMJ","JSYDGDMJ","PZSJ","SFYS","JD",
        //                         "ZT","PZSJXYYS","SJCJSJ","SJMJ","SBMJ","CZ","DB"};
        //string[] types = { "LONG", "TEXT", "TEXT", "LONG", "LONG", "LONG", "LONG", "LONG", "LONG",
        //                         "DATE", "DATE","TEXT","LONG", "LONG", "LONG", "DATE","TEXT","TEXT",
        //                     "TEXT","TEXT","DOUBLE","DOUBLE","DOUBLE","LONG"};
        XmlDocument xmlDoc;
        DataTable dt = new DataTable();


        
        public FormMain()
        {
            InitializeComponent();
            this.Text = "Excel2Shape  V" + Assembly.GetExecutingAssembly().GetName().Version.ToString();
            Control.CheckForIllegalCrossThreadCalls = false;


            XmlNodeList nodeList = XMLHelper.GetXmlNodeListByXpath(Path.Combine(Application.StartupPath, "Fields.xml"), "//fieldlist//field");
            dt.Columns.Add("fields", typeof(string));
            dt.Columns.Add("types", typeof(string));
            dt.Columns.Add("length", typeof(int));

            foreach (XmlNode xn in nodeList)
            {
               XmlElement xe = (XmlElement)xn;
               DataRow dr = dt.NewRow();
               dr["fields"] = xe.GetAttribute("name");
               dr["types"] = xe.GetAttribute("type");
               if (xe.GetAttribute("length") != "")
               {
                   dr["length"] =Convert.ToInt32(xe.GetAttribute("length"));
               }
              
               
               dt.Rows.Add(dr);
            }
        }





        private void button1_Click(object sender, EventArgs e)
        {
            
            using (BackgroundWorker bw = new BackgroundWorker())
            {
                bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
                bw.DoWork += new DoWorkEventHandler(bw_DoWork);
                bw.RunWorkerAsync("Tank");
            }       
            

        }

        void disableUI()
        {
            textBox1.Enabled = false;
            textBox2.Enabled = false;
            txbX.Enabled = false;
            txbY.Enabled = false;
            btnSetInput.Enabled = false;
            btnSetOutput.Enabled = false;
            btnTransform.Enabled = false;
            radioToPoint.Enabled = false;
            radioToLine.Enabled = false;
            radioToPloygon.Enabled = false;
        }
        void enableUI()
        {
            textBox1.Enabled = true;
            //textBox1.Text = "";
            textBox2.Enabled = true;
            //textBox2.Text = "";
            txbX.Enabled = true;
            txbY.Enabled = true;
            btnSetInput.Enabled = true;
            btnSetOutput.Enabled = true;
            btnTransform.Enabled = true;
            radioToPoint.Enabled = true;
            radioToLine.Enabled = true;
            radioToPloygon.Enabled = true;
        }
        bool checkSetting()
        {
            if (string.IsNullOrWhiteSpace(textBox1.Text))
            {
                if (MessageBox.Show("请选择输入路径！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                { }
                return false;
            }
            if (string.IsNullOrWhiteSpace(textBox2.Text))
            {
                if (MessageBox.Show("请选择输出路径！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                { }
                return false;
            }
            string inputPath = textBox1.Text.ToString();
            string outputPath = textBox2.Text.ToString();
            if (inputPath == outputPath)
            {
                if (MessageBox.Show("输出路径请勿与输出路径相同！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                { }
                return false;
            }
            if (txbX.Text.ToString() == "")
            {
                if (MessageBox.Show("横坐标不能为空！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                { }
                return false;
            }
            if (txbY.Text.ToString() == "")
            {
                if (MessageBox.Show("纵坐标不能为空！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                { }
                return false;
            }
            if (!Directory.Exists(inputPath))
            {
                if (MessageBox.Show("输入路径不存在！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                { }
                return false;
            }
            return true;
        }
        void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            disableUI();
            if (!checkSetting()) return;

            string inputPath = textBox1.Text.ToString();
            string outputPath = textBox2.Text.ToString();
            //文件夹及子文件夹下的所有文件的全路径
            string[] files = Directory.GetFiles(inputPath, "*.xls", SearchOption.AllDirectories);
            if (files.Length == 0)
            {
                if (MessageBox.Show("不存在Excel文件！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                { }
                return;
            }

            string[] outfiles = new string[files.Length];
            for (int i = 0; i < files.Length; i++)
            {
                string str =Path.GetDirectoryName(files[i].Substring(inputPath.Length, files[i].Length - inputPath.Length));
                outfiles[i] = str + "\\" + Path.GetFileNameWithoutExtension(files[i]);//只取文件名
            }

            int errorCount = 0;
            List<string> errorFiles = new List<string>();
            progressBar1.Maximum = files.Length;
            lblProgress.Text = "进度："+ progressBar1.Value.ToString() + "/" + files.Length.ToString()+"  正在初始化...";
            for (int i = 0; i < files.Length; i++)
            {
                try
                {
                    TransForm(files[i], outputPath + outfiles[i] + ".shp", outputPath);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message);
                    errorCount++;
                    errorFiles.Add(files[i]);
                }
                finally
                {
                    progressBar1.Value = i + 1;
                    lblProgress.Text = "进度：" + progressBar1.Value.ToString() + "/" + files.Length.ToString() + "   转换：" + files[i];
                }
            }

            string mess = "完成！共转换" + files.Length.ToString() + "个，失败" + errorCount.ToString() + "个！";
            if (MessageBox.Show(mess, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
            { progressBar1.Value = 0; lblProgress.Text = ""; }


            string errFilePath = Path.Combine(outputPath, "Error");
            //string errFilePath = outputPath + "\\log";
            if (!Directory.Exists(errFilePath))
            {
                Directory.CreateDirectory(errFilePath);
            }
            else
            {
                Directory.Delete(errFilePath, true);
                Directory.CreateDirectory(errFilePath);
            }
            //写日志
            System.IO.FileStream fs = new System.IO.FileStream(Path.Combine(outputPath, "log.txt"), FileMode.Append);
            StreamWriter sw = new StreamWriter(fs,Encoding.Default);
            //转换失败的写入日志
            sw.WriteLine(System.DateTime.Now.ToString());
            sw.WriteLine("转化失败(已复制至Error文件夹下)："+ errorCount.ToString());

            for (int i = 0; i < errorCount; i++)
			{
			    sw.WriteLine(errorFiles[i]);
                string copytoFile = errFilePath + errorFiles[i].Substring(inputPath.Length, errorFiles[i].Length - inputPath.Length);
                if (!Directory.Exists(Path.GetDirectoryName(copytoFile)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(copytoFile));
                }
                File.Copy(errorFiles[i], copytoFile , true);
			}

            //关闭日志文件
            sw.Close();
            fs.Close();

            //删除临时文件
            string[] tempfiles = Directory.GetFiles(outputPath, "temp.*");
            foreach (string str in tempfiles)
            {
                File.Delete(str);
            }

        }

        void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //这时后台线程已经完成，并返回了主线程，所以可以直接使用UI控件了 
            enableUI();
        }

        private void TransForm(string Inputfile,string Output,string outputPath)
        {
            //构造Geoprocessor  
            Geoprocessor gp = new Geoprocessor();
            string path = Path.GetDirectoryName(Output);
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            try
            {
                ESRI.ArcGIS.DataManagementTools.MakeXYEventLayer step1 = new ESRI.ArcGIS.DataManagementTools.MakeXYEventLayer();
                step1.table = Inputfile + "\\Sheet1$";//textBox1.Text.ToString() + "\\xy.xls\\Sheet1$";
                step1.in_x_field = txbX.Text.ToString();
                step1.in_y_field = txbY.Text.ToString();
                RunTool(gp, step1, null);
                Console.WriteLine(Inputfile);
                //Console.WriteLine("MakeXYEventLayer：添加点数据");

                ESRI.ArcGIS.DataManagementTools.PointsToLine step2 = new ESRI.ArcGIS.DataManagementTools.PointsToLine();
                step2.Input_Features = step1.out_layer;
                step2.Output_Feature_Class = outputPath + "\\temp.shp";//@"C:\Users\Yayu.Jiang\Desktop\output\temp.shp";
                RunTool(gp, step2, null);
                //Console.WriteLine("PointsToLine：点转线");

                ESRI.ArcGIS.DataManagementTools.FeatureToPolygon step3 = new ESRI.ArcGIS.DataManagementTools.FeatureToPolygon();
                step3.in_features = step2.Output_Feature_Class;
                step3.out_feature_class = Output;//@"C:\Users\Yayu.Jiang\Desktop\output\xy.shp";
                RunTool(gp, step3, null);
                //Console.WriteLine("FeatureToPolygon：要素转面");

                ESRI.ArcGIS.DataManagementTools.CalculateField calculateField = new ESRI.ArcGIS.DataManagementTools.CalculateField();
                calculateField.in_table = Output;
                calculateField.field = "ID";
                //calculateField.expression = "5";
                calculateField.expression = Path.GetFileNameWithoutExtension(Output);
                calculateField.expression_type = "VB";
                RunTool(gp, calculateField, null);
                //Console.WriteLine("CalculateField：填充ID字段");

                ESRI.ArcGIS.DataManagementTools.AddField addField = new ESRI.ArcGIS.DataManagementTools.AddField();
                addField.in_table = Output;
                addField.field_name = "YSBH";
                addField.field_type = "TEXT";
                addField.field_length = 50;
                RunTool(gp, addField, null);
                //Console.WriteLine("AddField：添加字段YSBH");

                string[] foldName = Inputfile.Split('\\');
                calculateField.in_table = addField.out_table;
                calculateField.field = "YSBH";
                calculateField.expression_type = "VB";
                calculateField.expression = "\"" + foldName[foldName.Length - 2] + "\"";
                //calculateField.expression = "\"aaa\"";
                RunTool(gp, calculateField, null);
                //Console.WriteLine("CalculateField：填充字段YSBH");

                #region 添加字段
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow dr = dt.Rows[i];
                    string[] fields = { dr["fields"].ToString(), dr["types"].ToString(), dr["length"].ToString() };
                    addField.field_name = fields[0];
                    addField.field_type = fields[1];
                    if (fields[1] == "TEXT")
                        addField.field_length = Convert.ToInt32(fields[2]);
                    RunTool(gp, addField, null);
                    //Console.WriteLine("AddField：添加字段" + dr["fields"].ToString());
                }

                #endregion

            }
            catch(Exception ex)
            {
                throw ex;
            }

        }

        private void RunTool(Geoprocessor geoprocessor, IGPProcess process, ITrackCancel TC)
        {
            // Set the overwrite output option to true  
            geoprocessor.OverwriteOutput = true;

            try
            {
                geoprocessor.Execute(process, null);
                ReturnMessages(geoprocessor);

            }
            catch (Exception err)
            {
                Console.WriteLine(err.Message);
                ReturnMessages(geoprocessor);
                throw err;
            }
        }  
   
        //Function for returning the tool messages.  
        private void ReturnMessages(Geoprocessor gp)
        {
            string ms = "";
            if (gp.MessageCount > 0)
            {
                for (int Count = 0; Count <= gp.MessageCount - 1; Count++)
                {
                    ms += gp.GetMessage(Count);
                }
            }
        }


        private List<string> correctExcel(string[] strFileNames)
        {
            List<string> correctFiles = new List<string>();
            bool isCopy = false;
            object missing = System.Reflection.Missing.Value;
            Excel.Application excel = new Excel.ApplicationClass();//lauch excel application
            if (excel == null)
            {
                new Exception("未能打开Excel文件！");
            }

            for (int index = 0; index < strFileNames.Length; index++)
            {

                excel.Visible = false; excel.UserControl = true;
                // 打开EXCEL文件
                Excel.Workbook wb = null;
                Excel.Worksheet ws = null;
                try
                {
                    wb = excel.Application.Workbooks.Open(strFileNames[index], missing, missing, missing, missing, missing,
                     missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    //取得第一个工作薄
                    ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);     
                }
                catch
                {
                    progressBar1.Value = index + 1;
                    lblProgress.Text = "进度：" + progressBar1.Value.ToString() + "/" + strFileNames.Length.ToString() + "   检验：" + strFileNames[index];
                    continue;
                }
                //取得总记录行数    (包括标题列)
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
                int columnsint = ws.UsedRange.Cells.Columns.Count;//得到列数
                //int columnsint = mySheet.UsedRange.Cells.Columns.Count;
                //获取第一行和最后一行数据   (不包括标题列) 
                char endColumn = (char)((int)'A' + columnsint - 1);
                Excel.Range rngStart = ws.Cells.get_Range("A2", endColumn.ToString() + "2");
                Excel.Range rngEnd = ws.Cells.get_Range("A" + rowsint, endColumn.ToString() + rowsint);

                object[,] arry1 = (object[,])rngStart.Value2;   //get range's value  
                object[,] arry2 = (object[,])rngEnd.Value2;

                if (arry1 == null || arry2 == null)
                {
                    progressBar1.Value = index + 1;
                    lblProgress.Text = "进度：" + progressBar1.Value.ToString() + "/" + strFileNames.Length.ToString() + "   检验：" + strFileNames[index];
                    continue;
                }
                for (int i = 3; i <= arry1.Length; i++)
                {
                    if (arry1[1, i].Equals(arry2[1, i])) continue;
                    else
                    {
                        isCopy = true;
                        break;
                    }

                }
                if (isCopy)
                {
                    for (int i = 1; i <= arry1.Length; i++)
                    {
                        ws.Cells[rowsint + 1, i] = arry1[1, i];
                    }
                    wb.Save();
                    correctFiles.Add(strFileNames[index]);//纪录修正的文件
                }
                wb.Close();
                
                progressBar1.Value = index + 1;
                lblProgress.Text = "进度：" + progressBar1.Value.ToString() + "/" + strFileNames.Length.ToString() + "   检验：" + strFileNames[index];

            }

            excel.Quit(); excel = null;
            Process[] procs = Process.GetProcessesByName("excel");
            foreach (Process pro in procs)
            {
                pro.Kill();//没有更好的方法,只有杀掉进程
            }
            GC.Collect();
            return correctFiles;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            setInput();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            setOutput();
        }

        private void setInput()
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = folderBrowserDialog1.SelectedPath;
            }
        }
        private void setOutput()
        {
            if (this.folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                this.textBox2.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void 设置属性ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmFields nfrmField = new frmFields(dt);
            nfrmField.Show();
 
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            string xmlfile = Path.Combine(Application.StartupPath, "Fields.xml");

            XmlNodeList xnl = XMLHelper.GetXmlNodeListByXpath(xmlfile, "//fieldlist//field");
            for (int i = 0; i < xnl.Count; i++)
            {
                XMLHelper.DeleteXmlNodeByXPath(xmlfile, "//fieldlist//field");
            }


            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                string [] xmlAttributeNames = {"name","type","length"};
                string [] values = {dr["fields"].ToString(),dr["types"].ToString(),dr["length"].ToString()};
                XMLHelper.CreateXmlNodeByXPath(xmlfile, "//fieldlist", "field", "", xmlAttributeNames, values);
            }

            //清理excel进程
            Process[] procs = Process.GetProcessesByName("excel");
            foreach (Process pro in procs)
            {
                pro.Kill();//没有更好的方法,只有杀掉进程
            }
            GC.Collect();
            
        }

        private void 设置输入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setOutput();
        }

        private void 设置输出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            setOutput();
        }

        private void 打开输出文件夹ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string outputPath = textBox2.Text.ToString();
            if (outputPath == "")
            {
                MessageBox.Show("请指定输出路径！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!Directory.Exists(outputPath))
            {
                MessageBox.Show("路径不存在！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            System.Diagnostics.Process.Start("explorer.exe", outputPath);
        }

        private void 打开日志ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string logfile = Path.Combine(textBox2.Text.ToString(),"log.txt");
            if (!File.Exists(logfile))
            {
                MessageBox.Show("暂无日志文件！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            System.Diagnostics.Process.Start("explorer.exe", logfile);
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox frmabout = new AboutBox();
            frmabout.ShowDialog();
        }


        #region 数据预处理部分
        private void 数据预处理ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (BackgroundWorker per = new BackgroundWorker())
            {
                per.RunWorkerCompleted += new RunWorkerCompletedEventHandler(per_RunWorkerCompleted);
                per.DoWork += new DoWorkEventHandler(per_DoWork);
                per.RunWorkerAsync("Tank");
            }      
        }
        void per_DoWork(object sender, DoWorkEventArgs e)
        {
            disableUI();

            if (!checkSetting()) return;

            string inputPath = textBox1.Text.ToString();
            string outputPath = textBox2.Text.ToString();
            //文件夹及子文件夹下的所有文件的全路径
            string[] files = Directory.GetFiles(inputPath, "*.xls", SearchOption.AllDirectories);
            if (files.Length == 0)
            {
                if (MessageBox.Show("不存在Excel文件！", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning) == DialogResult.OK)
                { }
                return;
            }

            progressBar1.Maximum = files.Length;
            lblProgress.Text = "预处理：" + progressBar1.Value.ToString() + "/" + files.Length.ToString() + "  正在初始化...";

            List<string> correctFiles = correctExcel(files);

            string mess = "完成！共处理" + files.Length.ToString() + "个，修正" + correctFiles.Count.ToString() + "个！";
            if (MessageBox.Show(mess, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
            { progressBar1.Value = 0; lblProgress.Text = ""; }

            //写日志
            System.IO.FileStream fs = new System.IO.FileStream(Path.Combine(outputPath, "log.txt"), FileMode.Append);
            StreamWriter sw = new StreamWriter(fs, Encoding.Default);
            //转换失败的写入日志
            sw.WriteLine(System.DateTime.Now.ToString());
            sw.WriteLine("修正多边形不闭合：" + correctFiles.Count.ToString());

            for (int i = 0; i < correctFiles.Count; i++)
            {
                sw.WriteLine(correctFiles[i]);
            }

            //关闭日志文件
            sw.Close();
            fs.Close();

        }

        void per_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //这时后台线程已经完成，并返回了主线程，所以可以直接使用UI控件了 
            enableUI();
        }

        #endregion


    }
}
