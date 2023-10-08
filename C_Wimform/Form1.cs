using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using System.Diagnostics;
using System.Data.SQLite;

namespace C_Wimform
{
    public partial class Form1 : Form
    {
        public Form1()
        {

            if (File.Exists(DDR_Record_str) == false || Directory.Exists(Repair_picture) == false /*|| Directory.Exists(@"D:\G947\5D.db") == false*/)
            {
                MessageBox.Show($"確認檔案是否存在於下列路徑\n{DDR_Record_str}\n{Repair_picture}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //跳錯誤訊息
            }
            else
            {
                InitializeComponent();
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyValue.ToString() == "13")  //按下Enter跑到textbox2
            {
                textBox2.Focus();
            }
        }
        ///////////////////////////全域變數 檔案存取位置///////////////////////////
        string Machine_name, Machine_Location,Qty;
        //string DDR_File_str = @"D:\5D_總表+排程+PT出貨.xlsx";
        //string DDR_File_str_copy = @"D:\5D_總表+排程+PT出貨_copy0.xlsx";
        //string DDR_Table_str = @"D:\DDR_table.xlsx";
        //string DDR_Table_str_copy = @"D:\DDR_table_copy0.xlsx";
        //string DDR_Table_Display = @"D:\DDR_table.xlsx";
        //string DDR_Table_Display_copy = @"D:\DDR_table_display0.xlsx";
        string DDR_Record_str = @"D:\G947\OSE TS不良品紀錄-DDR5.xlsx";
        string Repair_picture = @"D:\G947\Repair_picture\";
        string _connectionString = @"Data Source=D:\G947\5D.db";
        //string _connectionString = @"Data Source=G:\CC1 測試工程二課\H612( 謝宗達)\DDR_TOOL\5D.db";
        //string DDR_Record_str = @"G:\CC1 測試工程二課\H612( 謝宗達)\DDR_TOOL\OSE TS不良品紀錄-DDR5.xlsx";
        //string Repair_picture = @"G:\CC1 測試工程二課\H612( 謝宗達)\DDR_TOOL\Repair_picture\";
        Form2 form2 = new Form2();
        /////////////////////////////////////////////////////////////////////

        void Combo_Clear()  //清除所有資訊
        {
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            comboBox1.Text = "";
            comboBox1.Items.Clear();
            comboBox4.Text = "";
            comboBox4.Items.Clear();
            comboBox6.Text = "";
            comboBox6.Items.Clear();
            comboBox8.Text = "";
            comboBox8.Items.Clear();
            label7.Text = "";
            label10.Text = "";
            label13.Text = "";
            label16.Text = "";
        }


        public void textBox2_KeyDown(object sender, KeyEventArgs e) //工單按Enter
        {

            if (e.KeyValue.ToString() == "13")
            {
                Machine_name = "";
                Combo_Clear();
                form2.Hide();
                
                if (textBox2.Text != "")
                {
                    bool token = false;

                    using (var connection = new SQLiteConnection(_connectionString))
                    {
                        connection.Open();
                        var command = connection.CreateCommand();
                        command.CommandText = $@"SELECT * FROM Factory5D_all_data where wo='{textBox2.Text}'";

                        using (var reader=command.ExecuteReader())
                        {
                            if (reader.Read()==true)
                            {
                                textBox3.Text = reader.GetString(1);
                                textBox8.Text = reader.GetInt32(2).ToString();
                                Machine_Location = reader.GetString(3);
                                token = true;
                            }
                        }
                    }
                    if (token==true)
                    {
                        Machine_name = textBox3.Text.Substring(3, 4); //擷取機種文字
                    }


                    using (var connection = new SQLiteConnection(_connectionString))
                    {
                        connection.Open();
                        var command = connection.CreateCommand();
                        command.CommandText = $@"SELECT * FROM DDR_table where Machine_Name='{Machine_name}'";
                        string[] aaa ={"A0","B0","A1","B1" };

                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read() == true)
                            {
                                for (int i = 0; i < 4; i++)
                                {
                                    comboBox1.Items.Add(aaa[i]);
                                    comboBox4.Items.Add(aaa[i]);
                                    comboBox6.Items.Add(aaa[i]);
                                    comboBox8.Items.Add(aaa[i]);
                                }

                                string Repair_pit1 = $@"{Repair_picture}{Machine_name}.png";
                                if (File.Exists(Repair_pit1) == true)
                                {
                                    form2.testbox(Repair_pit1);
                                    form2.Show();
                                }
                                else
                                {
                                    MessageBox.Show("無此機種圖片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //跳錯誤訊息
                                }
                            }
                            else
                            {
                                MessageBox.Show("無此機種資訊", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //跳錯誤訊息
                            }
                        }
                    }
                    label19.Text = "";
                }
                else
                {
                    MessageBox.Show("請輸入工單", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //跳錯誤訊息
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox4.Text = "";
            label7.Text = "";
        }
        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox5.Text = "";
            label10.Text = "";
        }
        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox6.Text = "";
            label13.Text = "";
        }
        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox7.Text = "";
            label16.Text = "";
        }

        //////////////////////////////textbox自行輸入後按Enter/////////////////////////////////////////////////////////////////
        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            TextEnter(sender, e);
        }
        private void textBox5_KeyDown(object sender, KeyEventArgs e)
        {
            TextEnter(sender, e);
        }

        private void textBox6_KeyDown(object sender, KeyEventArgs e)
        {
            TextEnter(sender, e);
        }

        private void textBox7_KeyDown(object sender, KeyEventArgs e)
        {
            TextEnter(sender, e);
        }

        void TextEnter(object sender, KeyEventArgs e)
        {
            if (e.KeyValue.ToString() == "13")
            {
                label19.Text = "查詢中 請勿關閉應用程式";
                label7.Text = Combo_Select(comboBox1.Text, textBox4.Text);
                label10.Text = Combo_Select(comboBox4.Text, textBox5.Text);
                label13.Text = Combo_Select(comboBox6.Text, textBox6.Text);
                label16.Text = Combo_Select(comboBox8.Text, textBox7.Text);
            }
            label19.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DBsetting db = new DBsetting();
            db.Show();
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        string Combo_Select(string combo_num1, string combo_num2)
        {
            string comtext;
            string combotext;
            comtext = combo_num1 + combo_num2;

            try
            {
                using (var connection = new SQLiteConnection(_connectionString))
                {

                    connection.Open();
                    var command = connection.CreateCommand();
                    command.CommandText = $@"SELECT * FROM DDR_table where Machine_Name='{Machine_name}'";

                    using (var reader = command.ExecuteReader())
                    {
                        reader.Read();

                        if (comtext == "A00" || comtext == "A01" || comtext == "A02" || comtext == "A03" || comtext == "A04" || comtext == "A05" || comtext == "A06" || comtext == "A07")
                        {
                            return combotext = reader.GetString(1);
                        }
                        else if (comtext == "A08" || comtext == "A09" || comtext == "A010" || comtext == "A011" || comtext == "A012" || comtext == "A013" || comtext == "A014" || comtext == "A015")
                        {
                            return combotext = reader.GetString(2);
                        }
                        else if (comtext == "A016" || comtext == "A017" || comtext == "A018" || comtext == "A019" || comtext == "A020" || comtext == "A021" || comtext == "A022" || comtext == "A023")
                        {
                            return combotext = reader.GetString(3);
                        }
                        else if (comtext == "A024" || comtext == "A025" || comtext == "A026" || comtext == "A027" || comtext == "A028" || comtext == "A029" || comtext == "A030" || comtext == "A031")
                        {
                            return combotext = reader.GetString(4);
                        }
                        else if (comtext == "B00" || comtext == "B01" || comtext == "B02" || comtext == "B03" || comtext == "B04" || comtext == "B05" || comtext == "B06" || comtext == "B07")
                        {
                            return combotext = reader.GetString(5);
                        }
                        else if (comtext == "B08" || comtext == "B09" || comtext == "B010" || comtext == "B011" || comtext == "B012" || comtext == "B013" || comtext == "B014" || comtext == "B015")
                        {
                            return combotext = reader.GetString(6);
                        }
                        else if (comtext == "B016" || comtext == "B017" || comtext == "B018" || comtext == "B019" || comtext == "B020" || comtext == "B021" || comtext == "B022" || comtext == "B023")
                        {
                            return combotext = reader.GetString(7);
                        }
                        else if (comtext == "B024" || comtext == "B025" || comtext == "B026" || comtext == "B027" || comtext == "B028" || comtext == "B029" || comtext == "B030" || comtext == "B031")
                        {
                            return combotext = reader.GetString(8);
                        }
                        else if (comtext == "A10" || comtext == "A11" || comtext == "A12" || comtext == "A13" || comtext == "A14" || comtext == "A15" || comtext == "A16" || comtext == "A17")
                        {
                            return combotext = reader.GetString(9);
                        }
                        else if (comtext == "A18" || comtext == "A19" || comtext == "A110" || comtext == "A111" || comtext == "A112" || comtext == "A113" || comtext == "A114" || comtext == "A115")
                        {
                            return combotext = reader.GetString(10);
                        }
                        else if (comtext == "A116" || comtext == "A117" || comtext == "A118" || comtext == "A119" || comtext == "A120" || comtext == "A121" || comtext == "A122" || comtext == "A123")
                        {
                            return combotext = reader.GetString(11);
                        }
                        else if (comtext == "A124" || comtext == "A125" || comtext == "A126" || comtext == "A127" || comtext == "A128" || comtext == "A129" || comtext == "A130" || comtext == "A131")
                        {
                            return combotext = reader.GetString(12);
                        }
                        else if (comtext == "B10" || comtext == "B11" || comtext == "B12" || comtext == "B13" || comtext == "B14" || comtext == "B15" || comtext == "B16" || comtext == "B17")
                        {
                            return combotext = reader.GetString(13);
                        }
                        else if (comtext == "B18" || comtext == "B19" || comtext == "B110" || comtext == "B111" || comtext == "B112" || comtext == "B113" || comtext == "B114" || comtext == "B115")
                        {
                            return combotext = reader.GetString(14);
                        }
                        else if (comtext == "B116" || comtext == "B117" || comtext == "B118" || comtext == "B119" || comtext == "B120" || comtext == "B121" || comtext == "B122" || comtext == "B123")
                        {
                            return combotext = reader.GetString(15);
                        }
                        else if (comtext == "B124" || comtext == "B125" || comtext == "B126" || comtext == "B127" || comtext == "B128" || comtext == "B129" || comtext == "B130" || comtext == "B131")
                        {
                            return combotext = reader.GetString(16);
                        }
                        //else
                        //{
                        //    //Display_Save_Close();
                        //    //combotext = "無此位置";
                        //    //return combotext;
                        //    //MessageBox.Show("無此位置", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //跳錯誤訊息
                        //}
                    }
                }
            }
            catch
            {
                //Display_Save_Close();
                MessageBox.Show("無此位置", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //跳錯誤訊息
            }

            //void Display_Save_Close()
            //{

            //}
            return "";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            Qty = textBox8.Text;
            if (label10.Text == "" && label13.Text == "" && label16.Text == "" && textBox9.Text=="" && label7.Text == "" || textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("工號、工單、維修位置 勿空白", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);  //跳錯誤訊息
            }
            else
            {
                Excel.Application app_DDR_Record = new Excel.Application();
                Excel.Workbook wb1_DDR_Record = app_DDR_Record.Workbooks.Open(DDR_Record_str, Password:"E253");
                Excel.Worksheet ws1_DDR_Record = wb1_DDR_Record.Sheets[1];

                try
                {

                    string Employee_ID = textBox1.Text;
                    string Model_Name = textBox2.Text;
                    string Machine = textBox3.Text;
                    string Date = $"{DateTime.Now:yyyy-MM-dd}";
                    string combo1 = $"{comboBox1.Text}_{textBox4.Text}";
                    string combo2 = $"{comboBox4.Text}_{textBox5.Text}";
                    string combo3 = $"{comboBox6.Text}_{textBox6.Text}";
                    string combo4 = $"{comboBox8.Text}_{ textBox7.Text}";
                    string combo_ary = $"{combo1}/{combo2}/{combo3}/{combo4}";
                    int count = 0;
                    bool ws1DDR = true;

                        if (label7.Text != "")
                        {
                            count++;
                        }
                        if (label10.Text != "")
                        {
                            count++;
                        }
                        if (label13.Text != "")
                        {
                            count++;
                        }
                        if (label16.Text != "")
                        {
                            count++;
                        }

                        if (label7.Text == label10.Text || label7.Text == label13.Text || label7.Text == label16.Text)
                        {
                            label7.Text = "";
                            count--;
                        }
                        if (label10.Text == label13.Text || label10.Text == label16.Text)
                        {
                            if (label10.Text != "")
                            {
                                label10.Text = "";
                                count--;
                            }

                        }
                        if (label13.Text == label16.Text)
                        {
                            if (label13.Text != "")
                            {
                                label13.Text = "";
                                count--;
                            }

                        }
                        int total = count;

                    for (int i = 1; i <= ws1_DDR_Record.UsedRange.Rows.Count; i++)
                    {
                        if (Model_Name == ws1_DDR_Record.Cells[i, 2].Value.ToString() && combo_ary == ws1_DDR_Record.Cells[i, 8].Value.ToString() && Employee_ID == ws1_DDR_Record.Cells[i, 13].Value.ToString())
                        {
                            ws1_DDR_Record.Cells[i, 10].Value += 1;
                            ws1_DDR_Record.Cells[i, 6].Value = ws1_DDR_Record.Cells[i, 10].Value;
                            ws1_DDR_Record.Cells[i, 11].Value = ws1_DDR_Record.Cells[i, 10].Value * total;
                            ws1DDR = false;
                        }
                    }
                    if (ws1DDR == true)
                    {
                        string Error_info = $"{label7.Text}/{label10.Text}/{label13.Text}/{label16.Text}";

                        string[] ary = { "MB1", Model_Name, Machine, "", Qty, "1", Machine_Location, combo_ary, Error_info, "1", total.ToString(), "BDC02", Employee_ID, Date };
                        int Row_num = ws1_DDR_Record.UsedRange.Rows.Count;


                        for (int i = 0; i <= ary.GetUpperBound(0); i++)
                        {
                            ws1_DDR_Record.Cells[Row_num + 1, i + 1] = ary[i];
                        }
                    }

                    MessageBox.Show("儲存成功");
                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox9.Text = "";
                    comboBox1.Text = "";
                    comboBox4.Text = "";
                    comboBox6.Text = "";
                    comboBox8.Text = "";
                    label7.Text = "";
                    label10.Text = "";
                    label13.Text = "";
                    label16.Text = "";

                    wb1_DDR_Record.Save();
                    wb1_DDR_Record.Close();
                    app_DDR_Record.Quit();
                    Marshal.ReleaseComObject(app_DDR_Record);
                }
                catch (Exception)
                {
                    MessageBox.Show("儲存失敗\n確認檔案是否無法寫入");
                    wb1_DDR_Record.Save();
                    wb1_DDR_Record.Close();
                    app_DDR_Record.Quit();
                    Marshal.ReleaseComObject(app_DDR_Record);
                }
                
            }
        }


    }
}