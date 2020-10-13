using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;


namespace MSM_MACRO
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        Excel.Workbook wb = null;
        Excel.Worksheet ws = null;
        Excel.Application ap = null;
        Excel.Range range = null;
        string filePath;
        private void button2_Click(object sender, EventArgs e)
        {
            
            
            OpenFileDialog OFD = new OpenFileDialog();
            if (OFD.ShowDialog() == DialogResult.OK)
            {
                file_name.Text = OFD.FileName;
                filePath = OFD.FileName;
            }

            
            /*
            string str;
            int rCount;
            int cCount;
            int rw = 0;
            int cl = 0;
            */
            
            /*
            string a = ws.Cells[1, 1].value;
            label6.Text = a;
            */
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }


        string[] index_array = new string[1300];
        public void index_fx(int j)
        {
            if (checkBox3.Checked == true)
            {
                SendKeys.Send(add_word_fist.Text);
            }

            SendKeys.Send(index_array[j]);
            Thread.Sleep(70);
            
            if(checkBox2.Checked == true)
            {
                SendKeys.Send(add_word.Text);
            }
            
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(70);
        }

        public void backspace_fx(int j)
        {
            if (checkBox1.Checked == true)
            {
                for (int k = 0; k < index_array[j].Length + 3 + add_word.Text.Length + add_word_fist.Text.Length; k++)
                {
                    SendKeys.SendWait("{BACKSPACE}");
                    Thread.Sleep(20);
                }
            }
        }



       


        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        int cul_value;
        int row_value;

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor = new Cursor(Cursor.Current.Handle);

            int move_to_x = int.Parse(X_location.Text);
            int move_to_y = int.Parse(Y_location.Text);
            

            ap = new Excel.Application();
            wb = ap.Workbooks.Open(filePath);
            ws = wb.Worksheets[sheet_box.Text] as Excel.Worksheet;

            //string[] index_array = new string[1300];
            int row_range_total = int.Parse(row_range_end.Text)-int.Parse(row_range_start.Text) +1;
            row_value = int.Parse(row_range_start.Text);
            cul_value = int.Parse(cul_range.Text);
            
            for (int i=0; i<row_range_total; i++)
            {
                index_array[i] = ws.Cells[row_value, cul_value].value.ToString();
                row_value += 1;
            }

            Cursor.Position = new System.Drawing.Point(move_to_x, move_to_y);
            const uint LBUTTONDOWN = 0x0002;
            const uint LBUTTONUP = 0x0004;
            mouse_event(LBUTTONDOWN, 0, 0, 0, 0);
            mouse_event(LBUTTONUP, 0, 0, 0, 0);

            string dan_k1 = dan_1.Text;
            string dan_k2 = dan_2.Text;
            string dan_k3 = dan_3.Text;
            string dan_k4 = dan_4.Text;
            string dan_k5 = dan_5.Text;
            string dan_k6 = dan_6.Text;
            string dan_k7 = dan_7.Text;
            string dan_k8 = dan_8.Text;


            for (int j = 0; j < row_range_total; j++)
            {

                /*
                 public void run_fxx(string cb, string cb_text, string j, string input_dan){      //    cb_1, cb_2, cb_3    /  n , 인 , Del /  j / dan_k1, dan_k2
                    if (cb.Text != "")                    //"" 이거 공백 진짜 공백일 때 되는지 확인하고  세개 기능 확인하면 메서드화 시키셈
                    {                                       // asdf.runfx(1);  {cb_아몰랑}  public void asdf(속성 변수, 속성 변수) 해서 변수에 cb_1 , dan_1 이런식으로 넣어도 될듯
                        if (cb.Text.Equals(cb_text))
                        {
                            SendKeys.SendWait(input_dan);
                            Thread.Sleep(20);
                        }
                        else if (cb.Text == cb_text)
                        {
                            index_fx(j);
                        }
                        else if (cb.Text == cb_text)
                        {
                            backspace_fx(j);
                        }
                    }
                }
                  
                 
                 * */
                if (cb_1.Text != "")                    //"" 이거 공백 진짜 공백일 때 되는지 확인하고  세개 기능 확인하면 메서드화 시키셈
                {                                       // asdf.runfx(1);  {cb_아몰랑}  public void asdf(속성 변수, 속성 변수) 해서 변수에 cb_1 , dan_1 이런식으로 넣어도 될듯
                    if (cb_1.Text.Equals("1"))
                    {
                        SendKeys.SendWait(dan_k1);
                        Thread.Sleep(20);
                    }
                    else if (cb_1.Text == "인")
                    {
                        index_fx(j);
                    }
                    else if (cb_1.Text == "Del")
                    {
                        backspace_fx(j);
                    }
                }
                ////////////////////////////////////////////////////////
                if (cb_2.Text != "")                    //"" 이거 공백 진짜 공백일 때 되는지 확인하고  세개 기능 확인하면 메서드화 시키셈
                {                                       // asdf.runfx(1);  {cb_아몰랑}  public void asdf(속성 변수, 속성 변수) 해서 변수에 cb_1 , dan_1 이런식으로 넣어도 될듯
                    if (cb_2.Text.Equals("2"))
                    {
                        SendKeys.SendWait(dan_k2);
                        Thread.Sleep(20);
                    }
                    else if (cb_2.Text == "인")
                    {
                        index_fx(j);
                    }
                    else if (cb_2.Text == "Del")
                    {
                        backspace_fx(j);
                    }
                }
                if (cb_3.Text != "")                    //"" 이거 공백 진짜 공백일 때 되는지 확인하고  세개 기능 확인하면 메서드화 시키셈
                {                                       // asdf.runfx(1);  {cb_아몰랑}  public void asdf(속성 변수, 속성 변수) 해서 변수에 cb_1 , dan_1 이런식으로 넣어도 될듯
                    if (cb_3.Text.Equals("3"))
                    {
                        SendKeys.SendWait(dan_k3);
                        Thread.Sleep(20);
                    }
                    else if (cb_3.Text == "인")
                    {
                        index_fx(j);
                    }
                    else if (cb_3.Text == "Del")
                    {
                        backspace_fx(j);
                    }
                }

                if (cb_4.Text != "")                    //"" 이거 공백 진짜 공백일 때 되는지 확인하고  세개 기능 확인하면 메서드화 시키셈
                {                                       // asdf.runfx(1);  {cb_아몰랑}  public void asdf(속성 변수, 속성 변수) 해서 변수에 cb_1 , dan_1 이런식으로 넣어도 될듯
                    if (cb_4.Text.Equals("4"))
                    {
                        SendKeys.SendWait(dan_k4);
                        Thread.Sleep(20);
                    }
                    else if (cb_4.Text == "인")
                    {
                        index_fx(j);
                    }
                    else if (cb_4.Text == "Del")
                    {
                        backspace_fx(j);
                    }
                }

                if (cb_5.Text != "")                    //"" 이거 공백 진짜 공백일 때 되는지 확인하고  세개 기능 확인하면 메서드화 시키셈
                {                                       // asdf.runfx(1);  {cb_아몰랑}  public void asdf(속성 변수, 속성 변수) 해서 변수에 cb_1 , dan_1 이런식으로 넣어도 될듯
                    if (cb_5.Text.Equals("5"))
                    {
                        SendKeys.SendWait(dan_k5);
                        Thread.Sleep(20);
                    }
                    else if (cb_5.Text == "인")
                    {
                        index_fx(j);
                    }
                    else if (cb_5.Text == "Del")
                    {
                        backspace_fx(j);
                    }
                }
                if (cb_6.Text != "")                    //"" 이거 공백 진짜 공백일 때 되는지 확인하고  세개 기능 확인하면 메서드화 시키셈
                {                                       // asdf.runfx(1);  {cb_아몰랑}  public void asdf(속성 변수, 속성 변수) 해서 변수에 cb_1 , dan_1 이런식으로 넣어도 될듯
                    if (cb_6.Text.Equals("6"))
                    {
                        SendKeys.SendWait(dan_k6);
                        Thread.Sleep(20);
                    }
                    else if (cb_6.Text == "인")
                    {
                        index_fx(j);
                    }
                    else if (cb_6.Text == "Del")
                    {
                        backspace_fx(j);
                    }
                }

                if (cb_7.Text != "")                    //"" 이거 공백 진짜 공백일 때 되는지 확인하고  세개 기능 확인하면 메서드화 시키셈
                {                                       // asdf.runfx(1);  {cb_아몰랑}  public void asdf(속성 변수, 속성 변수) 해서 변수에 cb_1 , dan_1 이런식으로 넣어도 될듯
                    if (cb_7.Text.Equals("7"))
                    {
                        SendKeys.SendWait(dan_k7);
                        Thread.Sleep(20);
                    }
                    else if (cb_7.Text == "인")
                    {
                        index_fx(j);
                    }
                    else if (cb_7.Text == "Del")
                    {
                        backspace_fx(j);
                    }
                }

                if (cb_8.Text != "")                    //"" 이거 공백 진짜 공백일 때 되는지 확인하고  세개 기능 확인하면 메서드화 시키셈
                {                                       // asdf.runfx(1);  {cb_아몰랑}  public void asdf(속성 변수, 속성 변수) 해서 변수에 cb_1 , dan_1 이런식으로 넣어도 될듯
                    if (cb_8.Text.Equals("8"))
                    {
                        SendKeys.SendWait(dan_k8);
                        Thread.Sleep(20);
                    }
                    else if (cb_8.Text == "인")
                    {
                        index_fx(j);
                    }
                    else if (cb_8.Text == "Del")
                    {
                        backspace_fx(j);
                    }
                }









                //////////////////////////////////////////////////////
                //SendKeys.Send(dan_k1);
                //Thread.Sleep(70);

                /*
                if (checkBox1.Checked == true) { 
                    for (int k = 0; k < index_array[j].Length + 3 + add_word.Text.Length; k++)
                    {
                        SendKeys.SendWait("{BACKSPACE}");
                        Thread.Sleep(20);
                    }
                }
                */


                /*
                SendKeys.Send(index_array[j]);
                Thread.Sleep(70);
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(70);
                */



            }



        }

        private void button1_Click_1(object sender, EventArgs e)
        {

             

            


        }

        private void label18_Click(object sender, EventArgs e)
        {
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            int x_locate = MousePosition.X;
            int y_locate = MousePosition.Y;

            label18.Text = x_locate.ToString();
            label19.Text = y_locate.ToString();
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void label19_Click(object sender, EventArgs e)
        {

        }

        private void number_box_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_1.Text == "인") { 
                dan_1.Text = "------";
                textBox1.Text = "INDEX 입력";
            }
            if (cb_2.Text == "인")
            {
                dan_2.Text = "------";
                textBox3.Text = "INDEX 입력";
            }
            if (cb_3.Text == "인")
            {
                dan_3.Text = "------";
                textBox5.Text = "INDEX 입력";
            }
            if (cb_4.Text == "인")
            {
                dan_4.Text = "------";
                textBox7.Text = "INDEX 입력";
            }
            if (cb_5.Text == "인")
            {
                dan_5.Text = "------";
                textBox9.Text = "INDEX 입력";
            }
            if (cb_6.Text == "인")
            {
                dan_6.Text = "------";
                textBox11.Text = "INDEX 입력";
            }
            if (cb_7.Text == "인")
            {
                dan_7.Text = "------";
                textBox13.Text = "INDEX 입력";
            }
            if (cb_8.Text == "인")
            {
                dan_8.Text = "------";
                textBox15.Text = "INDEX 입력";
            }
            //////////////////////////////////////
            if (cb_1.Text == "Del")
            {
                dan_1.Text = "------";
                textBox1.Text = "INDEX 지우기";
            }
            if (cb_2.Text == "Del")
            {
                dan_2.Text = "------";
                textBox3.Text = "INDEX 지우기";
            }
            if (cb_3.Text == "Del")
            {
                dan_3.Text = "------";
                textBox5.Text = "INDEX 지우기";
            }
            if (cb_4.Text == "Del")
            {
                dan_4.Text = "------";
                textBox7.Text = "INDEX 지우기";
            }
            if (cb_5.Text == "Del")
            {
                dan_5.Text = "------";
                textBox9.Text = "INDEX 지우기";
            }
            if (cb_6.Text == "Del")
            {
                dan_6.Text = "------";
                textBox11.Text = "INDEX 지우기";
            }
            if (cb_7.Text == "Del")
            {
                dan_7.Text = "------";
                textBox13.Text = "INDEX 지우기";
            }
            if (cb_8.Text == "Del")
            {
                dan_8.Text = "------";
                textBox15.Text = "INDEX 지우기";
            }
        }

        private void Clear_Btn_Click(object sender, EventArgs e)
        {
            cb_1.Text = "";
            cb_2.Text = "";
            cb_3.Text = "";
            cb_4.Text = "";
            cb_5.Text = "";
            cb_6.Text = "";
            cb_7.Text = "";
            cb_8.Text = "";
            dan_1.Text = "";
            dan_2.Text = "";
            dan_3.Text = "";
            dan_4.Text = "";
            dan_5.Text = "";
            dan_6.Text = "";
            dan_7.Text = "";
            dan_8.Text = "";
            textBox1.Text = "";
            textBox3.Text = "";
            textBox5.Text = "";
            textBox7.Text = "";
            textBox9.Text = "";
            textBox11.Text = "";
            textBox13.Text = "";
            textBox15.Text = "";
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
