using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {
        Excel.Application exApp;
        Excel._Worksheet workSheet;
        double[,] D;
        double[] R;
        int N = 8;
        Random r = new Random();

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double[] X = Enumerable.Range(1, N).Select(x => x * 1.0).ToArray();
            double[] Y = Enumerable.Range(1, N).Select(i => Math.Round(i + r.NextDouble(), 2)).ToArray();

            exApp = new Excel.Application();
            exApp.Visible = true;
            exApp.Workbooks.Add();

            chart1.ChartAreas[0].AxisY.Interval = 1;
            chart1.ChartAreas[0].AxisX.Maximum = N + 2;
            chart1.ChartAreas[0].AxisY.Maximum = 10;

            chart1.Series[0].Points.Clear();
            for (int i = 0; i < N; i++)
            {
                chart1.Series[0].Points.AddXY(X[i] + 1, Y[i] + 1);
            }

            workSheet = (Excel.Worksheet)exApp.ActiveSheet;

            workSheet.Cells[2, "B"] = "X";
            workSheet.Cells[3, "B"] = "Y";

            for (int i = 3; i <= X.Length + 1; i++)
            {
                workSheet.Cells[2, i] = X[i - 3];
                workSheet.Cells[3, i] = Y[i - 3];
            }

            D = new double[3, 3];
            R = new double[3];

            D[0, 0] = X.Sum(x => x * x * x * x);
            D[1, 0] = D[0, 1] = X.Sum(x => x * x * x);
            D[2, 0] = D[1, 1] = D[0, 2] = X.Sum(x => x * x);
            D[2, 1] = D[1, 2] = X.Sum(x => x);
            D[2, 2] = N;

            R[0] = Enumerable.Range(0, N).Select(i => X[i] * X[i] * Y[i]).Sum();
            R[1] = Enumerable.Range(0, N).Select(i => X[i] * Y[i]).Sum();
            R[2] = Y.Sum(y => y);

            workSheet.Cells[5, "C"] = "Массив D";
            workSheet.Cells[5, "H"] = "R";
            workSheet.Cells[5, "F"] = "Дельта";

            for (int i = 6; i < 9; i++)
            {
                for (int j = 2; j < 5; j++)
                {
                    workSheet.Cells[i, j] = D[i - 6, j - 2];
                }
                workSheet.Cells[i, "H"] = R[i - 6];
            }

            DArr(10, Met(0, D));
            DArr(14, Met(1, D));
            DArr(18, Met(2, D));

            Excel.Range rng = workSheet.Range["F6"];
            rng.Formula = "=MDETERM(B6:D8)";

            rng = workSheet.Range["F10"];
            rng.Formula = "=MDETERM(B10:D12)";

            rng = workSheet.Range["F14"];
            rng.Formula = "=MDETERM(B14:D16)";

            rng = workSheet.Range["F18"];
            rng.Formula = "=MDETERM(B18:D20)";

            rng = workSheet.Range["I10"];
            workSheet.Cells[10, "H"] = "a = ";
            rng.Formula = "=F10/F6";

            rng = workSheet.Range["I14"];
            workSheet.Cells[14, "H"] = "b = ";
            rng.Formula = "=F14/F6";

            rng = workSheet.Range["I18"];
            workSheet.Cells[18, "H"] = "c = ";
            rng.Formula = "=F18/F6";

            double a = Math.Round(double.Parse(workSheet.Cells[10, "I"].Text), 2);
            double b = Math.Round(double.Parse(workSheet.Cells[14, "I"].Text), 2);
            double c = Math.Round(double.Parse(workSheet.Cells[18, "I"].Text), 2);

            chart1.Series[1].Points.Clear();
            chart1.Series[1].Name = a.ToString("##.##") + "x^2 + " + b.ToString("##.##") + "x + " + c.ToString("##.##");

            for (int i = 0; i < N; i++)
            {
                chart1.Series[1].Points.AddXY(X[i], a * X[i] * X[i] + b * X[i] + c);
            }

            exApp.Windows[1].Close(false, false);
            exApp.Quit();
        }

        public void DArr(int a, double[,] mass)
        {
            for (int i = a; i < a + 3; i++)
            {
                for (int j = 2; j < 5; j++)
                {
                    workSheet.Cells[i, j] = mass[i - a, j - 2];
                }
            }
        }

        public double[,] Met(int b, double[,] mass)
        {
            double[,] A = new double[3, 3];
            for (int i = 0; i < 3; i++)
            {
                for (int j = 0; j < 3; j++)
                {
                    A[i, j] = mass[i, j];
                }
            }

            for (int i = 0; i < 3; i++)
            {
                A[i, b] = R[i];
            }
            return A;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            chart1.Series[0].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            chart1.Series[1].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Spline;
        }
    }
}
