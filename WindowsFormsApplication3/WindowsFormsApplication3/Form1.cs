using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {
        private Excel.Application exApp;
        private Excel.Range exRng;
        private Excel.Sheets exSheets;
        private Excel.Worksheet  exWsheet;
        private Excel.Workbook exBook;
        private Excel.Workbooks exBooks;
        Random r = new Random();
        double[] mass = new double[6];

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            exApp = new Excel.Application();
            exApp.Visible = true;
            exApp.SheetsInNewWorkbook = 1;
            exApp.Workbooks.Add(Type.Missing);
            Excel._Worksheet workSheet = (Excel.Worksheet)exApp.ActiveSheet;

            for (int i = 0; i < 6; i++)
            {
                mass[i] = r.NextDouble() * 30;
                mass.Orderby = 
            }
            

            for (int i = 1; i < 7; i++)
			{
                workSheet.Cells[2, "B"] = "X";
                workSheet.Cells[3, "B"] = "Y";
                workSheet.Cells[2, i + 2] = i;
                workSheet.Cells[3, i + 2] = r.NextDouble() * 30;
                chart1.Series[0].Points.AddXY(xValue: i, yValue: r.NextDouble() * 30);
			}
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            exApp.Windows[1].Close(false, Type.Missing, Type.Missing);
            exApp.Quit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
        }
    }
}
