using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;

//97943
namespace Watchdog
{
    public partial class Form1 : Form
    {
        public delegate void AddListItem();
        public AddListItem myDelegate;

        Chart mych = new Chart();
        public Form1()
        {
            myDelegate = new AddListItem(makeGraph);
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Random rnd = new Random();           
            Color color = Color.DarkGray;
            mych.Series.Clear();

            var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                Name = "duck",
                Color = System.Drawing.Color.Green,
                IsVisibleInLegend = false,
                IsXValueIndexed = true,
                ChartType = SeriesChartType.Line
            };

            mych.Series.Add(series1);
            mych.Series["duck"].SetDefault(true);
            mych.Series["duck"].Enabled = true;
            mych.Series["duck"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            mych.Visible = true;
            mych.BackColor = color;
            mych.Titles.Add("test Value vs time");
            ChartArea chA = new ChartArea();
            chA.AxisX.Title = "Time in months";
            chA.AxisY.Title = "Value in us";
            chA.AxisX.Interval = 1;

            mych.ChartAreas.Add(chA);
            mych.Location = new Point(10, 10);
            mych.Width = 800;
            mych.Height = 500;
            for (int x = 0; x <= 12; x++)
            {
                int first = x;
                int second = rnd.Next(0, 50);
                mych.Series["duck"].Points.AddXY(first, second);
            }
            mych.Invalidate();
            mych.Show();
            Controls.Add(mych);
            vals = new List<double>();
        }

        public void makeGraph()
        {
            String title = Title;
            double[] value = vals.ToArray();
            Random rnd = new Random();
            Color color = Color.DarkGray;
            mych.Series.Clear();
            //mych.ChartAreas.Clear();

            var series1 = new System.Windows.Forms.DataVisualization.Charting.Series
            {
                Name = "duck",
                Color = System.Drawing.Color.Green,
                IsVisibleInLegend = false,
                IsXValueIndexed = true,
                ChartType = SeriesChartType.Line
            };

            mych.Series.Add(series1);
            mych.Series["duck"].SetDefault(true);
            mych.Series["duck"].Enabled = true;
            mych.Series["duck"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            mych.BackColor = color;
            mych.Titles.Add(title+"\nValue vs time");

            mych.Location = new Point(10, 10);
            mych.Width = 800;
            mych.Height = 500;
            for (int x = 0; x <= value.Length-1; x++)
            {
                int first = x;
                double second = value[x];
                mych.Series["duck"].Points.AddXY(first, second);
            }
            mych.Invalidate();
            mych.Show();
        }

        string Title = ""; List<double> vals;
        //Reads stock info from excel or csv.
        private void getStocks(int i)
        {
            //3112;
            int colCount = 2;
            Excel.Application xlApp = new Excel.Application();
            xlWorkbook1 = xlApp.Workbooks.Open(@"D:\Watchdog\Watchdog\Data\list.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook1.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            int Count = i;
            OpenBooks();
            for (int j = 1; j <= colCount; j++)
            {
                if(j==1)
                {
                    Console.Write("\n"+"||"+Count+"||");
                    Count++;
                }
                //write the value to the console
                if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                {
                    string value = xlRange.Cells[i, j].Value2.ToString();
                    Console.Write( value + "\t"); 
                    if(j==1 && i>1)
                    {                            
                        string price = getCurrentPrice(value);
                        savePrice(value, price,Week,i);
                    }
                } 
            }
             int row = 4;
            int rowCount = xlRange1.Rows.Count;
            int ColCount = xlRange1.Columns.Count;
            for (int col=1;col<= ColCount; col++)
            {
                if(col == 1)
                {
                    Title = xlRange1.Cells[row, col].Value2.ToString();
                }
                else
                {
                    if (xlRange1.Cells[row, col] != null && xlRange1.Cells[row, col].Value2 != null)
                    {
                        string str = xlRange1.Cells[row, col].Value2.ToString();
                        double s = 0;
                        double.TryParse(str, out s);
                        if (s > 0)
                        {
                            vals.Add(s);
                        }
                    }
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            CloseBooks();
            Marshal.ReleaseComObject(xlWorkbook1);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        int Week =2;

        private string getCurrentPrice(String url)
        {
            string price = "";
            System.Net.WebClient wc = new System.Net.WebClient();
            byte[] raw = wc.DownloadData("http://www.eoddata.com/stockquote/NYSE/"+url+".htm");

            string webData = System.Text.Encoding.UTF8.GetString(raw);
            webData = webData.Substring(webData.IndexOf("CONTENT_BEGIN"));
            webData = webData.Substring(webData.IndexOf("LAST:"));
            webData = webData.Substring(webData.IndexOf("26px")+12);
            price = webData.Substring(0, webData.IndexOf("<"));
            return price;
        }

        Excel.Application xlApp1 = new Excel.Application();
        Excel.Range xlRange1;
        Excel.Workbook xlWorkbook1;

        private void OpenBooks()
        {
            Excel._Worksheet xlWorksheet = xlWorkbook1.Sheets[2];
            xlRange1 = xlWorksheet.UsedRange;
        }

        private void CloseBooks()
        {            
            xlWorkbook1.Save();
            xlWorkbook1.Close();
            Console.Out.WriteLine("\n Saved and closed Excel");            
        }

        private void savePrice(string symb, string price,int x,int y)
        {
            Console.Out.WriteLine("Cell: " + "$"+price);//xlRange1.Cells[2, 1].Value2);
            xlRange1.Cells[y, x].Value = price;
            xlWorkbook1.Save();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Clockwork.Enabled = true;
        }

        private void Form1_FormClosing(object sender, EventArgs e)
        {
            try
            {
                xlWorkbook1.Close();
            }
            catch(Exception)
            {
                CloseBooks();
                Console.Out.WriteLine("Had problems while closing");
            }
         }

        int index = 1;

        private void button1_Click(object sender, EventArgs e)
        {
            //https://www.google.com/search?q=fun

            System.Net.WebClient wc = new System.Net.WebClient();
            byte[] raw = wc.DownloadData(@"https://www.google.com/search?q=fun");
            string webData = System.Text.Encoding.UTF8.GetString(raw);
            Console.WriteLine(webData);
        }

        private void Clockwork_Tick(object sender, EventArgs e)
        {
            Thread t = new Thread(() => getStocks(index));
            t.Start();
            if(index== 5)//how many stocks to look at
            {
                Week++;
                //vals = new List<double>();
                index = 0;
            }
            index++;
            if (index % 2 == 0)
            {
                if (vals.Count > 0)
                {
                    makeGraph();
                }
            }
        }

        private Thread myThread;
        private void ThreadFunction()
        {
            myThread = new Thread(new ThreadStart(ThreadFunctions));
            myThread.Start();
        }

        private void ThreadFunctions()
        {
            MyThreadClass myThreadClassObject = new MyThreadClass(this);
            myThreadClassObject.Run();
        }

        private int WebScore(String StockName)
        {
            //google search stock
            //take first 3 websites
            //point value sites based on good and bad phrases
            return 1;
        }

        private void Drawgraph(String StockName)
        {

        }
    }
    public class MyThreadClass
    {
        Form1 myFormControl1;
        public MyThreadClass(Form1 myForm)
        {
            
            myFormControl1 = myForm;
        }

        public void Run()
        {
            // Execute the specified delegate on the thread that owns
            // 'myFormControl1' control's underlying window handle.
            myFormControl1.Invoke(myFormControl1.myDelegate);
        }
    }
}
