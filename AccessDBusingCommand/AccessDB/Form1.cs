using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace AccessDB
{
    public partial class Form1 : Form
    {
        BindingSource BindingSource1;
        public Form1()
        {
            InitializeComponent();
            BindingSource1 = new BindingSource();
        }

        private void btnAccess_Click(object sender, EventArgs e)
        {
            //第一步：設定連線字串
            String strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=xml.accdb";

            //第二步：建立資料庫連線物件
            OleDbConnection cn = new OleDbConnection(strConnectionString);

            //第三步：開啟資料庫連線
            cn.Open();

            //第四步：取得資料並填入 DataSet
            OleDbCommand oledbCommand = new OleDbCommand("SELECT * FROM xml", cn);
            //oledbCommand.CommandText = "Text";


            //OleDbDataReader reader;

            DataSet ds = new DataSet();
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(oledbCommand))
            {
                adapter.Fill(ds);
            }

            /*
            OleDbDataAdapter oleda = new OleDbDataAdapter("SELECT * FROM 客戶", cn);
            DataSet ds = new DataSet("ds_客戶");
            oleda.Fill(ds, "ds_客戶");*/

            //第五步：設定 DataSource
            BindingSource1.DataSource = ds.Tables[0];
            DataGridView1.DataSource = BindingSource1;
            BindingNavigator1.BindingSource = BindingSource1;

            //第六步：關閉資料庫連線
            cn.Close();
        }

        private void btnAccessInsert_Click(object sender, EventArgs e)
        {
            //第一步：設定連線字串
            String strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=xml.accdb";

            //第二步：建立資料庫連線物件
            OleDbConnection cn = new OleDbConnection(strConnectionString);

            //第三步：開啟資料庫連線
            cn.Open();
            try
            {
                String strSQL = "INSERT INTO xml(識別碼,updatetime,Info) VALUES('" + TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "')";

                OleDbCommand cmdLiming = new OleDbCommand(strSQL, cn);
                cmdLiming.ExecuteNonQuery();

                //第五步：取得資料並填入 DataSet
                OleDbDataAdapter oleda = new OleDbDataAdapter("SELECT * FROM xml", cn);
                DataSet ds = new DataSet("ds_xml");
                oleda.Fill(ds, "ds_xml");

                //第六步：設定 DataSource
                BindingSource1.DataSource = ds.Tables[0];
                DataGridView1.DataSource = BindingSource1;
                BindingNavigator1.BindingSource = BindingSource1;

                //第七步：關閉資料庫連線
                cn.Close();

                MessageBox.Show("Insert successfully");
                DataGridView1.CurrentCell = DataGridView1.Rows[DataGridView1.RowCount - 2].Cells[0];
                DataGridView1.Rows[DataGridView1.CurrentRow.Index].Selected = true;
            }
            catch { 
                MessageBox.Show("重號辣幹!");
            }
            //★ 第四步：利用 OleDbCommand 執行 SQL 語法
           
        }

        private void btnAccessUpdate_Click(object sender, EventArgs e)
        {
            //第一步：設定連線字串
            String strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=xml.accdb";

            //第二步：建立資料庫連線物件
            OleDbConnection cn = new OleDbConnection(strConnectionString);

            //第三步：開啟資料庫連線
            cn.Open();

            //★ 第四步：利用 OleDbCommand 執行 SQL 語法
            String strSQL = "UPDATE xml SET Info='" + txtUpdate_Contact.Text + "', updatetime='" + txtUpdate_Name.Text + "' WHERE 識別碼='" + txtUpdate_NO.Text + "'";

            OleDbCommand cmdLiming = new OleDbCommand(strSQL, cn);
            cmdLiming.ExecuteNonQuery();

            //第五步：取得資料並填入 DataSet
            OleDbDataAdapter oleda = new OleDbDataAdapter("SELECT * FROM xml", cn);
            DataSet ds = new DataSet("ds_xml");
            oleda.Fill(ds, "ds_xml");

            //第六步：設定 DataSource
            BindingSource1.DataSource = ds.Tables[0];
            DataGridView1.DataSource = BindingSource1;
            BindingNavigator1.BindingSource = BindingSource1;

            //第七步：關閉資料庫連線
            cn.Close();

            MessageBox.Show("Update successfully");
            DataGridView1.CurrentCell = DataGridView1.Rows[DataGridView1.RowCount - 2].Cells[0];
        
        }

        private void btnAccessDelete_Click(object sender, EventArgs e)
        {
            //第一步：設定連線字串
            String strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=xml.accdb";

            //第二步：建立資料庫連線物件
            OleDbConnection cn = new OleDbConnection(strConnectionString);

            //第三步：開啟資料庫連線
            cn.Open();

            //★ 第四步：利用 OleDbCommand 執行 SQL 語法
            String strSQL = "DELETE FROM xml WHERE 識別碼 ='" + txtDelete_NO.Text + "'";

            OleDbCommand cmdLiming = new OleDbCommand(strSQL, cn);
            cmdLiming.ExecuteNonQuery();
            //第五步：取得資料並填入 DataSet
            OleDbDataAdapter oleda = new OleDbDataAdapter("SELECT * FROM xml", cn);
            DataSet ds = new DataSet("ds_xml");
            oleda.Fill(ds, "ds_xml");

            //第六步：設定 DataSource
            BindingSource1.DataSource = ds.Tables[0];
            DataGridView1.DataSource = BindingSource1;
            BindingNavigator1.BindingSource = BindingSource1;

            //第七步：關閉資料庫連線
            cn.Close();

            MessageBox.Show("Delete successfully");
            DataGridView1.CurrentCell = DataGridView1.Rows[DataGridView1.RowCount - 2].Cells[0];
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                int selected_row = DataGridView1.CurrentRow.Index;
                txtUpdate_NO.Text = DataGridView1.Rows[selected_row].Cells[0].FormattedValue.ToString();
                txtUpdate_Name.Text = DataGridView1.Rows[selected_row].Cells[1].FormattedValue.ToString();
                txtUpdate_Contact.Text = DataGridView1.Rows[selected_row].Cells[2].FormattedValue.ToString();

                txtDelete_NO.Text = DataGridView1.Rows[selected_row].Cells[0].FormattedValue.ToString();
            }
            catch (Exception ee)
            {
                txtUpdate_NO.Text = "";
                txtUpdate_Name.Text = "";
                txtUpdate_Contact.Text = "";

                txtDelete_NO.Text = "";
            }
        }

        private void btnAccessQuery_Click(object sender, EventArgs e)
        {
            //第一步：設定連線字串
            String strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=xml.accdb";

            //第二步：建立資料庫連線物件
            OleDbConnection cn = new OleDbConnection(strConnectionString);

            //第三步：開啟資料庫連線
            cn.Open();

            //★ 第四步：利用 OleDbCommand 執行 SQL 語法
            String sql = "SELECT * FROM xml WHERE ";

            if (TextBox4.Text != "" || TextBox5.Text != "")
            {
                if (TextBox4.Text == "")
                { sql += "info like '" + TextBox5.Text + "'"; }
                else if (TextBox5.Text == "")
                { sql += "識別碼 like '" + TextBox4.Text + "'"; }
                else if (TextBox4.Text != "" && TextBox5.Text != "")
                {
                    sql += "識別碼 like '" + TextBox4.Text + "'";
                    sql += " and ";
                    sql += "info like '" + TextBox5.Text + "'";
                }
            }

            //第五步：取得資料並填入 DataSet
            OleDbCommand oledbCommand = new OleDbCommand(sql, cn);
            //oledbCommand.CommandText = "Text";


            //OleDbDataReader reader;

            DataSet ds = new DataSet();
            OleDbDataAdapter adapter = new OleDbDataAdapter(oledbCommand);
            adapter.Fill(ds);


            //不採用 DataAdapter 方式
            /*
            OleDbDataAdapter oleda = new OleDbDataAdapter(sql, cn);
            DataSet ds = new DataSet("ds_客戶");
            oleda.Fill(ds, "ds_客戶");*/

            //第六步：設定 DataSource
            BindingSource1.DataSource = ds.Tables[0];
            DataGridView1.DataSource = BindingSource1;
            BindingNavigator1.BindingSource = BindingSource1;

            //第七步：關閉資料庫連線
            cn.Close();
        }        
    }
}
