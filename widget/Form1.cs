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
using System.Configuration;
using System.Runtime.InteropServices;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace widget
{
    public partial class Form1 : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private int doba = ((Convert.ToInt32(ConfigurationManager.AppSettings["interval"])) * 60) * 1000;

        static string[] set1 = ConfigurationManager.AppSettings["excel1"].Split(',');
        static string[] set2 = ConfigurationManager.AppSettings["excel2"].Split(',');
        static string[] set3 = ConfigurationManager.AppSettings["excel3"].Split(',');
        static string[] set4 = ConfigurationManager.AppSettings["excel4"].Split(',');

        string[,] setPole = new string[,] 
        {
            {set1[0].ToString(), set1[1].ToString() , set1[2].ToString() , set1[3].ToString() , set1[4].ToString() },
            {set2[0].ToString(), set2[1].ToString() , set2[2].ToString() , set2[3].ToString() , set2[4].ToString() },
            {set3[0].ToString(), set3[1].ToString() , set3[2].ToString() , set3[3].ToString() , set3[4].ToString() },
            {set4[0].ToString(), set4[1].ToString() , set4[2].ToString() , set4[3].ToString() , set4[4].ToString() },
        };
        public Form1()
        {
            InitializeComponent();
            timer.Start();
            timer.Interval = doba;
        }

        private void Panel_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            NactiData();
        }

        private void NactiData()
        {
            for(int i = 0; i < 4; i++)
            {
                if(Path.GetExtension(setPole[i, 0]).ToString() == ".xls")
                {
                    CellXls bunkaXLS = new CellXls(setPole[i, 0], setPole[i, 1], Convert.ToInt32(setPole[i, 2]), Convert.ToInt32(setPole[i, 3]));
                    bunkaXLS.ReadExcelFile();

                    if (i == 0)
                    {
                        cell1.Text = bunkaXLS.FillCell();
                        lbl1Popis.Text = setPole[i, 4];
                    }                      
                    else if (i == 1)
                    {
                        cell2.Text = bunkaXLS.FillCell();
                        lbl2Popis.Text = setPole[i, 4];
                    }                      
                    else if (i == 2)
                    {
                        cell3.Text = bunkaXLS.FillCell();
                        lbl3Popis.Text = setPole[i, 4];
                    }                       
                    else if (i == 3)
                    {
                        cell4.Text = bunkaXLS.FillCell();
                        lbl4Popis.Text = setPole[i, 4];
                    }                      
                }
                else if(Path.GetExtension(setPole[i, 0]).ToString() == ".xlsx")
                {
                    CellXlsX bunkaXLSX = new CellXlsX(setPole[i, 0], setPole[i, 1], Convert.ToInt32(setPole[i, 2]), Convert.ToInt32(setPole[i, 3]));
                    bunkaXLSX.ReadExcelFile();

                    if (i == 0)
                    {
                        cell1.Text = bunkaXLSX.FillCell();
                        lbl1Popis.Text = setPole[i, 4];
                    }
                    else if (i == 1)
                    {
                        cell2.Text = bunkaXLSX.FillCell();
                        lbl2Popis.Text = setPole[i, 4];
                    }
                    else if (i == 2)
                    {
                        cell3.Text = bunkaXLSX.FillCell();
                        lbl3Popis.Text = setPole[i, 4];
                    }
                    else if (i == 3)
                    {
                        cell4.Text = bunkaXLSX.FillCell();
                        lbl4Popis.Text = setPole[i, 4];
                    }
                }
                else
                {
                    MessageBox.Show(@"Není zadán soubor typu Excel");
                }
            }
        }
        private void CloseWidget(object sender, EventArgs e)
        {
            this.Close();
        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            NactiData();
        }



        public event EventHandler LoadCompleted;
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            this.OnLoadCompleted(EventArgs.Empty);
        }
        protected virtual void OnLoadCompleted(EventArgs e)
        {
            var handler = LoadCompleted;
            if (handler != null)
                handler(this, e);
        }
        private void MainForm_Load(object sender, EventArgs e)
        {
            //Just for test, you can make a delay to simulate a time-consuming task
            //In a real application here you load your data and other settings
        }


    }
}
