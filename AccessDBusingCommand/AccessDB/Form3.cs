using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AccessDB
{
    public partial class Form3 : Form
    {     
        Form2 form2;
        BindingSource BindingSource1;
        public Form3()
        {
            InitializeComponent();    
             BindingSource1 = new BindingSource();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                String strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=xml.accdb";

                OleDbConnection cn = new OleDbConnection(strConnectionString);

                cn.Open();

                String strSQL = "INSERT INTO 使用者([員工編號],[帳號],[密碼]) VALUES('" + textBox1.Text + "','" + textBox2.Text + "','" + textBox3.Text + "')";
                OleDbCommand cmdLiming = new OleDbCommand(strSQL, cn);
                cmdLiming.ExecuteNonQuery();

                OleDbDataAdapter oleda = new OleDbDataAdapter("SELECT * FROM 使用者", cn);
                DataSet ds = new DataSet("ds_使用者");
                oleda.Fill(ds, "ds_使用者");
                cn.Close();

                MessageBox.Show("新的使用者");               
                this.Visible = false;              
            }
            catch 
            {
                MessageBox.Show("已註冊過");
            }
        }
    }
}
