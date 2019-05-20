using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using xl = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System.Text.RegularExpressions;
using System.Windows.Forms.DataVisualization.Charting;

namespace XLS
{
    public partial class Form1 : Form
    {
        int cR = 0;
        private List<DataItem> actualValues = new List<DataItem>();
        private List<DataItem> forecastValues = new List<DataItem>();
        private List<DataItem> estimateValues = new List<DataItem>();

        string stDateTextBox;
        string enDateTextBox;

        DateTime startDate;
        DateTime endDate;

        int indexStartDate = 0;
        int indexEndDate;
        int daysDiff;

        DateTime dt;
        decimal val;
        decimal sum;
        decimal rand;

        decimal deltaEstimateValues;

        public Form1()
        {
            InitializeComponent();
        }
        
        public List<DataItem> GetActualValues()
        {
            Microsoft.Office.Interop.Excel.Application exl = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = exl.Workbooks.Open("K:\\Users\\Ira\\Desktop\\wb.xlsx");
            Microsoft.Office.Interop.Excel.Worksheet ws = exl.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            xl.Range uR = ws.UsedRange;
            cR = uR.Rows.Count;
            actualValues = new List<DataItem>();
            DateTime dt;

            dt = ws.Cells[1, 1].Value;
            dt = dt.AddDays(-1);
            actualValues.Add(new DataItem(dt, 0, 0));
            sum = 0;
            for (int i = 1; i <= cR; i++) 
            {
                dt = ws.Cells[i, 1].Value;
                decimal val = Convert.ToDecimal(ws.Cells[i, 2].Value);
                sum += val;
                actualValues.Add(new DataItem(dt, val, sum));
            }
            foreach(DataItem i in actualValues)
            {
                Console.WriteLine("actualValues"+i.date.ToString() + "  " + i.ptdValue + "  " + i.sumValue);
                Console.WriteLine("==================");
            }

            wb.Close(true, Type.Missing, Type.Missing);
            exl.Quit();
            return actualValues;
        }
        public List<DataItem> GetForecastValues(List<DataItem> actualValues)
        {
            daysDiff = (endDate - actualValues[actualValues.Count - 1].date).Days;

            val = actualValues[actualValues.Count - 1].ptdValue;
            sum = actualValues[actualValues.Count - 1].sumValue;
            dt = actualValues[actualValues.Count - 1].date;
            forecastValues.Add(new DataItem(dt, val, sum));
            Random r = new Random();
            for (int i = 0; i < daysDiff; i++)
            {
                rand = Convert.ToDecimal(r.Next(80, 121) / 100.00);
                val = Math.Round((sum * rand / actualValues.Count),2);
                sum += val;
                dt = dt.AddDays(1);
                forecastValues.Add(new DataItem(dt, val, sum));
            }
            foreach (DataItem i in forecastValues)
            {
                Console.WriteLine("forecastValues"+i.date.ToString() + "  " + i.ptdValue + "  " + i.sumValue);
                Console.WriteLine("==================");
            }
            return forecastValues;
        }
        public List<DataItem> GetEstimateValues(int indexStartDate, int indexEndDate) {
            dt = startDate.AddDays(-1);
            val = 0;
            sum = 0;
            estimateValues.Add(new DataItem(dt, val, sum));

            deltaEstimateValues = actualValues[indexStartDate - 1].sumValue * daysDiff / 10;

            for (int k=indexStartDate;k<actualValues.Count;k++)
            {
                val = actualValues[k].ptdValue + deltaEstimateValues;
                sum += val;
                dt = actualValues[k].date;
                estimateValues.Add(new DataItem(dt, val, sum));
            }
            for(int k=1;k<forecastValues.Count;k++)
            {
                val = forecastValues[k].ptdValue + deltaEstimateValues;
                sum += val;
                dt = forecastValues[k].date;
                estimateValues.Add(new DataItem(dt, val, sum));
            }
            foreach (DataItem i in estimateValues)
            {
                Console.WriteLine("estimateValues" + i.date.ToString() + "  " + i.ptdValue + "  " + i.sumValue);
                Console.WriteLine("==================");
            }
            return estimateValues;
        }
  

        private void Button1_Click(object sender, EventArgs e)
        {
            
            stDateTextBox = textBox1.Text;
            enDateTextBox = textBox2.Text;
            if (stDateTextBox == "" || enDateTextBox == "") MessageBox.Show("Need to point Start date and End date.\nPlease, enter dates in fortam 'dd.mm.yyyy'.");
            else
            {
                if (actualValues.Count == 0) actualValues = GetActualValues(); // получаем список actualValues из файла Excel
                forecastValues.Clear();
                estimateValues.Clear();

                startDate = DateTime.Parse(stDateTextBox);
                endDate = DateTime.Parse(enDateTextBox);

                indexEndDate = actualValues.Count - 1;
                for (int i = 0; i < actualValues.Count; i++)
                {
                    if (actualValues[i].date == startDate) // определяем индексы в списке для startDate и endDate
                    {
                        indexStartDate = i;
                    }
                    if (actualValues[i].date == endDate)
                    {
                        indexEndDate = i;
                    }
                }
                if (startDate < actualValues[0].date || startDate > actualValues[actualValues.Count - 1].date)
                {
                    MessageBox.Show("Need to point Start date at interval from " + String.Format("{0: dd.MM.yyyy}", actualValues[1].date) + " to " + String.Format("{0: dd.MM.yyyy}", actualValues[actualValues.Count - 1].date));
                }
                else if (startDate > endDate)
                {
                    MessageBox.Show("End date must be later than Start date!");
                }
                else
                {
                    if (indexEndDate == actualValues.Count - 1 & endDate != actualValues[actualValues.Count - 1].date)
                    {
                        forecastValues = GetForecastValues(actualValues);         //создаем список forecastValues 
                    }
                    estimateValues = GetEstimateValues(indexStartDate, indexEndDate);   //создаем список estimateValues

                    chart1.Series.Clear();
                    Axis ax = new Axis();
                    ax.Title = "time [d]";
                    chart1.ChartAreas[0].AxisX = ax;
                    Axis ay = new Axis();
                    ay.Title = "Costs [$]";
                    chart1.ChartAreas[0].AxisY = ay;
                    DrawValues("ActualValues", Color.Red, actualValues);
                    DrawValues("ForecastValues", Color.Blue, forecastValues);
                    DrawValues("EstimateValues", Color.Green, estimateValues);
                    Console.WriteLine("=======================");
                }
            }
        }
        private void DrawValues(string name, Color color, List<DataItem> values)
        {
            chart1.Series.Add(name);
            chart1.Series[name].Color = color;
            chart1.Series[name].ChartType = SeriesChartType.Spline;
            if (radioButton1.Checked) chart1.Series[name].ChartType = SeriesChartType.Spline;
            else if (radioButton2.Checked) chart1.Series[name].ChartType = SeriesChartType.Line;
            else if (radioButton3.Checked) chart1.Series[name].ChartType = SeriesChartType.Column;
            else if (radioButton4.Checked) chart1.Series[name].ChartType = SeriesChartType.SplineArea;
            else if (radioButton5.Checked) chart1.Series[name].ChartType = SeriesChartType.Point;
            else if (radioButton6.Checked) chart1.Series[name].ChartType = SeriesChartType.StepLine;
            foreach (var value in values)
            {
                chart1.Series[name].Points.AddXY(value.date, value.sumValue);
            }
        }
        public void WriteForecast(List<DataItem> forecastValues)
        {
            Microsoft.Office.Interop.Excel.Application exl = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = exl.Workbooks.Add();
            Microsoft.Office.Interop.Excel.Worksheet ws = exl.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            ws.Cells.ClearContents();
            
            cR = 0;
            for (int k = 1; k < forecastValues.Count; k++)
            {
                cR += 1;
                ws.Cells[cR, 1].Value = forecastValues[k].date;
                ws.Cells[cR, 2].Value = forecastValues[k].ptdValue;
            }
            ws.Cells[1, 1].EntireColumn.Autofit();
            wb.Close(true, "K:\\Users\\Ira\\Desktop\\wb1.xlsx");
            exl.Quit();

        }
        private void Button2_Click(object sender, EventArgs e)
        {
            WriteForecast(forecastValues);
        }
    }
}
