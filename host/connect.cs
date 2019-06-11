using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace host
{
    public partial class connect : Form
    {
        public connect()
        {
            InitializeComponent();
            label2.Text = "正在连接中...";
            Thread.Sleep(5000);
            label2.Text = "正在连接中...";
                        
        }
        public delegate void TransfDelegate(String value);

 

        private void button1_Click(object sender, EventArgs e)
        {
            var f = new OpenFileDialog();
            f.Filter = "Excel|*.xlsx";
            //f.Multiselect = true; //多选            
            if (f.ShowDialog() == DialogResult.OK)
            {
                String filepath = f.FileName;//eg.G:\新建文件夹\新建文本文档.txt                
                String filename = f.SafeFileName;//eg.新建文本文档.txt      
                
            }
            
           
        }
    }
}
