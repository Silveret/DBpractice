using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AccessDB
{
    public partial class Form2 : Form
    {
        string target = "https://youtu.be/dQw4w9WgXcQ?t=43";
        Form1 form1;
        Form3 form3 = new Form3();
        string a, b;
        int aa, bb;
        public Form2()
        {
            InitializeComponent();
            form1 = new Form1();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            a = textBox1.Text;
            b = textBox2.Text;
            String strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=xml.accdb";

            //第二步：建立資料庫連線物件
            OleDbConnection cn = new OleDbConnection(strConnectionString);

            //第三步：開啟資料庫連線
            cn.Open();
            if ((new System.Text.RegularExpressions.Regex("^[\u4e00-\u9fa5a-zA-Z0-9\b]$")).IsMatch(textBox1.Text) || 
                (new System.Text.RegularExpressions.Regex("^[\u4e00-\u9fa5a-zA-Z0-9\b]$")).IsMatch(textBox2.Text))
            {
                try
                {
                    //★ 第四步：利用 OleDbCommand 執行 SQL 語法
                    String sql = "SELECT * FROM 使用者 WHERE ";

                    if (textBox1.Text != "" && textBox2.Text != "")
                    {
                        sql += "帳號 like '" + textBox1.Text + "'";
                        sql += " and ";
                        sql += "密碼 like '" + textBox2.Text + "'";
                    }

                    //第五步：取得資料並填入 DataSet
                    OleDbCommand oledbCommand = new OleDbCommand(sql, cn);
                    //oledbCommand.CommandText = "Text";
                    //OleDbDataReader reader;            
                    object obj = oledbCommand.ExecuteScalar();
                    cn.Close();
                    form1.Visible = true;
                    this.Visible = false;
                    //System.Diagnostics.Process.Start(target);
                    MessageBox.Show("管理者" + textBox1.Text + "你好");

                }
                catch
                {
                    MessageBox.Show("媽的還想騙阿");
                }
            }
            else
            {
                MessageBox.Show("媽的還想injection我阿");
            }
        }

        private void label3_Click(object sender, EventArgs e)
        {
            form3.Visible = true;
        }

    }
}

