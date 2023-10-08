using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;

namespace C_Wimform
{
    public partial class DBsetting : Form
    {
        public DBsetting()
        {
            InitializeComponent();
        }

        string _connectionString = @"Data Source=D:\G947\5D.db";

        private void button1_Click(object sender, EventArgs e)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                try
                {
                    connection.Open();
                    var serach = connection.CreateCommand();
                    var command = connection.CreateCommand();
                    command.CommandText = $@"insert into Factory5D_all_data VALUES('{textBox1.Text}','{textBox2.Text}',{textBox3.Text},'{textBox4.Text}')";
                    command.ExecuteNonQuery();
                    MessageBox.Show("總表建立成功");
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox1.Focus();
                }
                catch (Exception)
                {
                    MessageBox.Show("WO:"+textBox1.Text + "已存在");
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox1.Focus();
                }
                
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            using (var connection = new SQLiteConnection(_connectionString))
            {
                try
                {
                    connection.Open();
                    var serach = connection.CreateCommand();
                    var command = connection.CreateCommand();
                    command.CommandText = $@"insert into DDR_table VALUES('{textBox5.Text}','{textBox6.Text}','{textBox7.Text}','{textBox8.Text}','{textBox9.Text}','{textBox10.Text}','{textBox11.Text}','{textBox12.Text}','{textBox13.Text}','{textBox14.Text}','{textBox15.Text}','{textBox16.Text}','{textBox17.Text}','{textBox18.Text}','{textBox19.Text}','{textBox20.Text}','{textBox21.Text}')";
                    command.ExecuteNonQuery();
                    MessageBox.Show("機種建立成功"); textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox9.Text = "";
                    textBox10.Text = "";
                    textBox11.Text = "";
                    textBox12.Text = "";
                    textBox13.Text = "";
                    textBox14.Text = "";
                    textBox15.Text = "";
                    textBox16.Text = "";
                    textBox17.Text = "";
                    textBox18.Text = "";
                    textBox19.Text = "";
                    textBox20.Text = "";
                    textBox21.Text = "";
                    textBox5.Focus();
                }
                catch (Exception)
                {
                    MessageBox.Show("機種:" + textBox5.Text + "已存在");
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox9.Text = "";
                    textBox10.Text = "";
                    textBox11.Text = "";
                    textBox12.Text = "";
                    textBox13.Text = "";
                    textBox14.Text = "";
                    textBox15.Text = "";
                    textBox16.Text = "";
                    textBox17.Text = "";
                    textBox18.Text = "";
                    textBox19.Text = "";
                    textBox20.Text = "";
                    textBox21.Text = "";
                    textBox5.Focus();
                }
                
            }
        }
        

        private void textBox22_KeyDown(object sender, KeyEventArgs e)
        {
            if ((int)e.KeyCode==13)
            {
                if ("3613131" == textBox22.Text)
                {
                    tabControl1.Visible = true;
                    label22.Visible = false;
                    textBox22.Visible = false;
                }
                else
                {
                    MessageBox.Show("密碼錯誤");
                    textBox22.Text = "";
                    textBox22.Focus();
                }
            }
            
        }
    }
}
