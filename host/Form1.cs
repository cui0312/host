using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.IO.Ports;
using System.Data.OleDb;
using System.Threading;
using System.Windows.Forms.DataVisualization.Charting;

namespace host
{
  
    public partial class Form1 : Form
    {
        public int page=0;
        public String filepath;
        public int timeindex = 0;
        private SerialPort ComDevice = new SerialPort();

        private const int LOCAL_PORT = 1500;
        
        TcpListener listener = null;
        private IPAddress serverIP;
        private IPAddress localaddr1;
        private IPEndPoint serverFullAddr;
        private Socket sock;
        private Socket newSocket;
        Thread myThread = null;
     

        public Form1()
        {
            InitializeComponent();
            textBox16.Text = GetIpAddress();
            System.DateTime currentTime = new System.DateTime();
            int 年 = currentTime.Year;
            int 月 = currentTime.Month;
            int 日 = currentTime.Day;
            int 时 = currentTime.Hour;
            int 分 = currentTime.Minute;
            int 秒 = currentTime.Second;
            timer1.Enabled = false;
            string url = Application.StartupPath + "/HTMLPage1.html";
            webBrowser1.Url = new Uri(url);//指定url 
          

        }
        private void CloseForm(object sender, FormClosedEventArgs e)
        {
            System.Diagnostics.Process.GetCurrentProcess().Kill(); //关闭窗口后关闭后台进程
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            webBrowser1.ScriptErrorsSuppressed = true;
            string path = Path.Combine(Application.StartupPath, "HTMLPage1.html");
            webBrowser1.Navigate(path);           
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {
        }
        
        private string GetIpAddress()
        {
            string hostName = Dns.GetHostName();   //获取本机名
            //IPHostEntry localhost = Dns.GetHostByName(hostName);    //方法已过期，可以获取IPv4的地址
            IPHostEntry localhost = Dns.GetHostEntry(hostName);   //获取IPv6地址
            // System.Net.IPAddress[] addressList = Dns.GetHostAddresses(hostName);//会返回所有地址，包括IPv4和IPv6  
            localaddr1 = localhost.AddressList[0];
            return localaddr1.ToString();
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label21_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
           // myThread = new Thread(new ThreadStart(BeginListen));
           // myThread.Start();
            start.Enabled = false;
            end.Enabled = true;
            connect f = new connect();
            f.ShowDialog();
            f.Owner = this;
           
        }


        private void BeginListen()
        {
            serverIP = localaddr1;
            {
                if (tbxPort.Text == null)
                {
                    MessageBox.Show("请先输入端口号！");
                }
                else
                {
                    //serverFullAddr = new IPEndPoint(serverIP, int.Parse(tbxPort.Text));//设置IP，端口
                }
            }
           // sock = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            //sock.Bind(serverFullAddr);
           // sock.Listen(3);
            newSocket = sock.Accept();

            byte[] message = new byte[1024];
            string mess = "";
            while (true)
            {
                try
                {
                    int bytes = newSocket.Receive(message);
                    mess = Encoding.Default.GetString(message, 0, bytes);
                    Invoke(new PrintRecvMssgDelegate(PrintRecvMssg), new object[] { mess });
                    MessageBox.Show("建立成功");
                }
                catch (Exception ee)
                {
                    Invoke(new PrintRecvMssgDelegate(PrintRecvMssg1), new object[] { "建立连接出错" + ee });
                }
            }
        }
        private void PrintRecvMssg(string info)
        {
            Form newmess = new Form();
            newmess.ShowDialog();
           // RTXT.Text += string.Format("{0}\r\n", info);
        }
        private void PrintRecvMssg1(string info)
        {
            //label3.Text = string.Format("{0}\r\n", info);
        }



        private delegate void PrintRecvMssgDelegate(string s);

        private void chartrefresh_Click(object sender, EventArgs e)
        {
            pointrefresh();
        }
        private void pointrefresh ()

        {
            chart1.DataSource = GetData();
            // Set series members names for the X and Y values
            chart1.Series["uav1"].XValueMember = "X1";
            chart1.Series["uav1"].YValueMembers = "Y1";

            chart1.Series["uav2"].XValueMember = "X2";
            chart1.Series["uav2"].YValueMembers = "Y2";

            chart1.Series["uav3"].XValueMember = "X3";
            chart1.Series["uav3"].YValueMembers = "Y3";

            
            // Data bind to the selected data source
            chart1.DataBind();
          
            // Set series chart type
            // chart1.Series["Series1"].ChartType = SeriesChartType.Line;
            //chart1.Series["Series2"].ChartType = SeriesChartType.Spline;
            // Set point labels
            chart1.Series["uav1"].IsValueShownAsLabel = false;
            chart1.Series["uav2"].IsValueShownAsLabel = false;
            chart1.Series["uav3"].IsValueShownAsLabel = false;
            // Enable X axis margin
            chart1.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
            chart1.ChartAreas[0].AxisY.LabelStyle.Enabled = false;
            chart1.ChartAreas["ChartArea1"].AxisX.IsMarginVisible = true;
            chart1.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;//隐藏刻度线
            // Enable 3D, and show data point marker lines
            //chart1.ChartAreas["ChartArea1"].Area3DStyle.Enable3D = true;
            chart1.Series["uav1"]["ShowMarkerLines"] = "False";
            chart1.Series["uav2"]["ShowMarkerLines"] = "False";
            chart1.Series["uav3"]["ShowMarkerLines"] = "False";
            //chart1.Series["uav4"]["ShowMarkerLines"] = "False";
            //chart1.Series["uav5"]["ShowMarkerLines"] = "False";
            this.chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false; //Y轴的网格线去掉
            this.chart1.ChartAreas[0].AxisY.MajorGrid.Enabled = false; //Y轴的网格线去掉
        }


        private static DataSet dynamicpoint()
        {
            string filepath = "C:/hostdata.xlsx";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;data source=" + filepath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1'";
            DataSet ds = new DataSet();
            OleDbDataAdapter oada = new OleDbDataAdapter("select * from [Sheet1$]", strConn);
            oada.Fill(ds);         
            return ds;
        }
        public DataSet ds = dynamicpoint();
        private void dynamic(int timeindex)
        {
            Random random = new Random();
            string n = (random.Next(0, 1) + 5).ToString();
            if (page == 1)
            {
                uav1x.Text = ds.Tables[0].Rows[timeindex][0].ToString();
                uav1y.Text = ds.Tables[0].Rows[timeindex][2].ToString();
                v1.Text = ds.Tables[0].Rows[timeindex][4].ToString();
                psiv1.Text = ds.Tables[0].Rows[timeindex][5].ToString();
                uav1z.Text = n;
            }
            if (page == 2)
            {
                uav1x.Text = ds.Tables[0].Rows[timeindex][6].ToString();
                uav1y.Text = ds.Tables[0].Rows[timeindex][8].ToString();
                v1.Text = ds.Tables[0].Rows[timeindex][10].ToString();
                psiv1.Text = ds.Tables[0].Rows[timeindex][11].ToString();
                uav1z.Text = n;
            }

            if (page == 3)
            {
                uav1x.Text = ds.Tables[0].Rows[timeindex][12].ToString();
                uav1y.Text = ds.Tables[0].Rows[timeindex][14].ToString();
                v1.Text = ds.Tables[0].Rows[timeindex][16].ToString();
                psiv1.Text = ds.Tables[0].Rows[timeindex][17].ToString();
                uav1z.Text = n;
            }
            this.chart3.ChartAreas[0].AxisX.MajorGrid.Enabled = false; //Y轴的网格线去掉
            this.chart3.ChartAreas[0].AxisY.MajorGrid.Enabled = false; //Y轴的网格线去掉
            this.chart3.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;

            this.chart3.Series["uav1"].Points.AddXY(ds.Tables[0].Rows[timeindex][0], ds.Tables[0].Rows[timeindex][2]);
            this.chart3.Series["uav2"].Points.AddXY(ds.Tables[0].Rows[timeindex][6], ds.Tables[0].Rows[timeindex][8]);
            this.chart3.Series["uav3"].Points.AddXY(ds.Tables[0].Rows[timeindex][12], ds.Tables[0].Rows[timeindex][14]);
            //this.chart3.Series["uav4"].Points.AddXY(ds.Tables[0].Rows[timeindex][6], ds.Tables[0].Rows[timeindex][7]);
            //this.chart3.Series["uav5"].Points.AddXY(ds.Tables[0].Rows[timeindex][8], ds.Tables[0].Rows[timeindex][9]);
            chart3.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;//隐藏刻度线
            //chart3.ChartAreas["ChartArea1"].AxisX.IsMarginVisible = true;
            chart3.ChartAreas[0].AxisX.LabelStyle.Enabled = false;
            chart3.ChartAreas[0].AxisY.LabelStyle.Enabled = false;
            chart3.ChartAreas[0].AxisX.Enabled = AxisEnabled.False;
            chart3.ChartAreas[0].AxisY.Enabled = AxisEnabled.False;
        }
        public static DataTable ReadExcelToTable()//excel存放的路径  
      {
            
            //连接字符串
            string filepath = "C:/hostdata.xlsx";
            string strCon = "Provider=Microsoft.ACE.OLEDB.12.0;data source=" + filepath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1'";
            // Office 07及以上版本 不能出现多余的空格 而且分号注意  
            //string connstring = Provider=Microsoft.JET.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; //Office 07以下版本   
            using (OleDbConnection conn = new OleDbConnection(strCon))  
          {  
              conn.Open();
                //得到所有sheet的名字   
             DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
                //得到第一个sheet的名字    
               
             string firstSheetName = sheetsName.Rows[0][2].ToString();
             
                //查询字符串  
              string sql = string.Format("SELECT * FROM [{0}]", firstSheetName);              
             //string sql = string.Format("SELECT * FROM [{0}] WHERE [日期] is not null", firstSheetName); //查询字符串  
             OleDbDataAdapter ada = new OleDbDataAdapter(sql, strCon);  
             DataSet set = new DataSet();  
             ada.Fill(set);
             return set.Tables[0];  
        }
           
    }
        private DataTable GetData()
        {
            DataTable mydata = ReadExcelToTable();
            DataColumn dcX1 = mydata.Columns["X1"];
            DataColumn dcY1 = mydata.Columns["Y1"];
            DataColumn dcX2 = mydata.Columns["X2"];
            DataColumn dcY2 = mydata.Columns["Y2"];
            DataColumn dcX3 = mydata.Columns["X3"];
            DataColumn dcY3 = mydata.Columns["Y3"];
            //  DataColumn dcY1 = new DataColumn("Y1", Type.GetType("System.Int32"));
            return mydata;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (timeindex < 100)
            {
                timeindex++;
                timer3.Enabled = true;
                dynamic(timeindex);
            }
            else
            {
                timer1.Enabled = false;
                timeindex = 0;
        }

        }

        private void button1_Click_2(object sender, EventArgs e)
        {

            var f = new OpenFileDialog();
            f.Filter = "Excel|*.xlsx";
            //f.Multiselect = true; //多选            
            if (f.ShowDialog() == DialogResult.OK) {
                String filepath = f.FileName;//eg.G:\新建文件夹\新建文本文档.txt                
                String filename = f.SafeFileName;//eg.新建文本文档.txt                
                this.textBox20.Text = filename;
            }
                  
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
        }

        private void end_Click(object sender, EventArgs e)
        {
            try
            {
                newSocket.Close();
                sock.Close();
                myThread.Abort();
                start.Enabled = true;
                end.Enabled = false;
            }
            catch (Exception ee)
            {
                MessageBox.Show("未检测Client接入" + ee);
            }
        }

        private void 开始连接ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myThread = new Thread(new ThreadStart(BeginListen));
            myThread.Start();
            start.Enabled = false;
            end.Enabled = true;
        }

        private void 断开连接ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                newSocket.Close();
                sock.Close();
                myThread.Abort();
                start.Enabled = true;
                end.Enabled = false;
            }
            catch (Exception ee)
            {
                MessageBox.Show("未检测Client接入" + ee);
            }
        }

        private void 查看访问文件要求ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("请添加EXCEL文件！5组仿真数据按照列存储，第一行为列名。");
        }

        private void 添加仿真文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var f = new OpenFileDialog();
            f.Filter = "Excel|*.xlsx";
            //f.Multiselect = true; //多选            
            if (f.ShowDialog() == DialogResult.OK)
            {
                String filepath = f.FileName;//eg.G:\新建文件夹\新建文本文档.txt                
                String filename = f.SafeFileName;//eg.新建文本文档.txt                
                this.textBox20.Text = filename;
            }
        }

        private void 关闭仿真文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.textBox20.Text = "";
        }

        private void 连接设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form set = new Form();
            set.ShowDialog();

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            label24.Text = "系统时间："+DateTime.Now.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form dataform = new Form();

        }

        private void wbShow_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }
        
        private void button3_Click_1(object sender, EventArgs e)
        {
            timer1.Enabled = true;
           
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            chart3.Series["uav1"].Points.Clear();
            chart3.Series["uav2"].Points.Clear();
            chart3.Series["uav3"].Points.Clear();
            timer3.Enabled = false;
        }

        private void 查看无人机数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 软件信息ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            dr = MessageBox.Show("功能有待完善！");
        }

        private void 软件说明ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dr;
            dr = MessageBox.Show("功能有待完善！");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox20.Text = "C:/hostdata.xlsx";
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            page = 1;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            page = 2;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            page = 3;
        }
    }
    

}
