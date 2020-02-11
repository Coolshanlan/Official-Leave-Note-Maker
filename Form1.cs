using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.SS;
using NPOI.XWPF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
namespace 公假單製作
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        ComboBox[] NameS = new ComboBox[6];
        TextBox[] Number = new TextBox[6];
        TextBox[] Date = new TextBox[6];
        ComboBox[] Session = new ComboBox[6];
        List<s_d> student_list = new List<s_d>();
        int students = 0;

        void Read_Excel(string Excelstring)
        {
            // FileStream fs = new FileStream(@"資研社名單.xls", FileMode.Open, FileAccess.Read);
            FileStream fs = new FileStream(Excelstring+".xls", FileMode.Open, FileAccess.Read);
            IWorkbook work = new HSSFWorkbook(fs);
            ISheet sheet = work.GetSheetAt(0);
             student_list = new List<s_d>();
            for (int i = 1; i < sheet.LastRowNum+1 ; i++)
            {
                s_d sd = new s_d();
                sd.Class = sheet.GetRow(i).GetCell(0).StringCellValue;
                sd.Number = sheet.GetRow(i).GetCell(1).StringCellValue;
                sd.Name = sheet.GetRow(i).GetCell(2).StringCellValue;
                student_list.Add(sd);
            }
            fs.Close();
            work.Clear();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            
            document = app.Documents.Open(Path.GetFullPath("空白公假單.docx"));
            
            app.Visible = false;
            app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            ft = new Word.Font();
            ft.Name = "Arial";
            ft.Size =  14;
            
            int i = 0;
            for(int ii = 0; ii < 5; ii++)
            {
                NameS[ii] = new ComboBox();
                Number[ii] = new TextBox();
                Date[ii] = new TextBox();
                Session[ii] = new ComboBox();
            }
            NameS[i++] = comboBox7;
            NameS[i++] = comboBox8;
            NameS[i++] = comboBox9;
            NameS[i++] = comboBox10;
            NameS[i++] = comboBox11;
            NameS[i++] = comboBox13;
            i = 0;
            Number[i++] = textBox14;
            Number[i++] = textBox13;
            Number[i++] = textBox12;
            Number[i++] = textBox11;
            Number[i++] = textBox10;
            Number[i++] = textBox9;
            i = 0;
            Date[i++] = textBox3;
            Date[i++] = textBox4;
            Date[i++] = textBox5;
            Date[i++] = textBox6;
            Date[i++] = textBox7;
            Date[i++] = textBox8;
            i = 0;
            Session[i++] = comboBox18;
            Session[i++] = comboBox19;
            Session[i++] = comboBox20;
            Session[i++] = comboBox21;
            Session[i++] = comboBox22;
            Session[i++] = comboBox15;
            for(int ii = 0; ii < 6; ii++)
            {
                NameS[ii].Tag = ii.ToString();
                Number[ii].Tag = ii.ToString();
                Date[ii].Tag = ii.ToString();
                Session[ii].Tag = ii.ToString();
            }
            comboBox2.SelectedIndex = 0;

            comboBox1.Items.Clear();
        }
        Word.Application app = new Word.Application();
        Word.Document document;
        Word.Font ft;
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Write_word();
                document.SaveAs(Path.GetFullPath(textBox1.Text + ".doc"), Word.WdSaveFormat.wdFormatDocument97);
            }
            catch (Exception ex)
            {
                MessageBox.Show("檔名錯誤 \r\n\r\n " + ex.Message + "");
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox2.Text = System.DateTime.Now.Month.ToString("00") + "/" + System.DateTime.Now.Day.ToString("00");
            textBox1.Text = System.DateTime.Now.Month.ToString("00") + System.DateTime.Now.Day.ToString("00")+"公假單" ;
        }

        public void students_count()
        {
            students = 0;
            foreach (Control c in NameS)
            {
                if (c.Text != "") students++;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                app.Visible = false;
                app.Quit();
                // document = new Word.Document();
                document.Close();

            }
            catch
            {

            }
        }
        void Write_word()
        {
            students_count();
            for(int i = 2; i < 8; i++)
            {
                document.Tables[1].Cell(i, 1).Range.Font = ft;
                document.Tables[1].Cell(i, 2).Range.Font = ft;
                document.Tables[1].Cell(i, 3).Range.Font = ft;
                document.Tables[1].Cell(i, 4).Range.Font = ft;
                document.Tables[1].Cell(i, 5).Range.Font = ft;
                document.Tables[2].Cell(i, 1).Range.Font = ft;
                document.Tables[2].Cell(i, 2).Range.Font = ft;
                document.Tables[2].Cell(i, 3).Range.Font = ft;
                document.Tables[2].Cell(i, 4).Range.Font = ft;
                document.Tables[2].Cell(i, 5).Range.Font = ft;
                document.Tables[1].Cell(i, 1).Range.Text = comboBox1.Text;
                document.Tables[1].Cell(i, 2).Range.Text = Number[i - 2].Text;
                document.Tables[1].Cell(i, 3).Range.Text = NameS[i - 2].Text;
                document.Tables[1].Cell(i, 4).Range.Text = Date[i - 2].Text;
                document.Tables[1].Cell(i, 5).Range.Text = Session[i - 2].Text;
                document.Tables[2].Cell(i, 1).Range.Text = comboBox1.Text;
                document.Tables[2].Cell(i, 2).Range.Text = Number[i - 2].Text;
                document.Tables[2].Cell(i, 3).Range.Text = NameS[i - 2].Text;
                document.Tables[2].Cell(i, 4).Range.Text = Date[i - 2].Text;
                document.Tables[2].Cell(i, 5).Range.Text = Session[i - 2].Text;
                if (i > students+1 )
                {
                    document.Tables[1].Cell(i, 1).Range.Text = "";
                    document.Tables[2].Cell(i, 1).Range.Text = "";
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            
            Write_word();
            if(MessageBox.Show("確定列印？", "", MessageBoxButtons.OKCancel) == DialogResult.OK)
            {
                document.PrintOut();
            }
        }


        private void button4_Click(object sender, EventArgs e)
        {
            foreach(Control c in groupBox2.Controls)
            {
                if (c.Tag != null) c.Text = "";
            }
        }



        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                Write_word();
                document.SaveAs(Path.GetFullPath(textBox1.Text + ".pdf"),  (Word.WdSaveFormat.wdFormatPDF));
            }
            catch(Exception ex)
            {
                MessageBox.Show("檔名錯誤 \r\n\r\n "+ex.Message+"");
            }
        }
        private void combobox_items_update()
        {
            students_count();
            List<s_d> now_student = now_student_update(comboBox1.Text);

            for(int i = 0; i < 6; i++)
            {
                while(NameS[i].Items.Count > now_student.Count)
                {
                    NameS[i].Items.RemoveAt(NameS[i].Items.Count-1);
                }
            }
            if (students+1 < 6)
                for (int i = students+1; i < 6; i++)
                {
                    NameS[i].Items.Clear();
                }
            
            if (students != 0)NameS[students-1].Items.Add("");

        }
        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
            button4.PerformClick();
            List<s_d> now_student = now_student_update(comboBox1.Text);
            foreach(ComboBox c in NameS)
            {
                c.Items.Clear();
            }
            for(int i = 0; i < 6; i++)
            {
                foreach(var a in now_student)
                {
                    NameS[i].Items.Add(a.Name);
                }
                if(i < now_student.Count)
                NameS[i].Text = now_student[i].Name;
            }
            if(students+1<6)
            for(int i = students+1; i < 6; i++)
            {
                NameS[i].Items.Clear();
            }
            combobox_items_update();
            button3.PerformClick();
            comboBox12_TextChanged(null, null);
        }
        private List<s_d> now_student_update (string Class)
        {
            List<s_d> student_now = new List<s_d>();
            student_now = student_list.FindAll(x=>x.Class == Class);
            return student_now;
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            students_count();
            for(int i = 0; i < students; i++)
            {
                Date[i].Text = textBox2.Text;
            }
        }
        bool clear_b = false;
        private void comboBox7_TextChanged(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            List<s_d> now_student = now_student_update(comboBox1.Text);
            int cb_index = int.Parse(cb.Tag.ToString());
            if(cb.Text == "")
            {
                Date[cb_index].Text = "";
                Session[cb_index].Text = "";
                Number[cb_index].Text = "";
            }
            else
            {
                var aa = student_list.FindAll(x=>x.Name == cb.Text);
                if (aa.Count != 0) Number[cb_index].Text = aa[0].Number;
                Date[cb_index].Text = textBox2.Text;
                Session[cb_index].Text = comboBox12.Text;
                if(students+1< 6)
                {
                    foreach (var a in now_student)
                    {
                        NameS[students+1].Items.Add(a.Name);
                    }
                }
            }
            combobox_items_update();
        }

        private void comboBox12_TextChanged(object sender, EventArgs e)
        {
            students_count();
            for (int i = 0; i < students; i++)
            {
                Session[i].Text = comboBox12.Text;
            }
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            Read_Excel(comboBox2.Text);
            var aa = student_list.GroupBy(x => x.Class).Select(x => x.ToList()).ToList();
            foreach (var a in aa)
            {
                comboBox1.Items.Add(a[0].Class);
            }
            comboBox1.SelectedIndex = 3;
        }
    }
    class s_d
    {
        public string Name { get; set; }
        public string Number { get; set; }
        public string Class { get; set; }
    }
}
