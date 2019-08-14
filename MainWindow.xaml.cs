using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace report_generator_LERS
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        bool flag_day = Properties.Settings.Default.dayType;
        Worksheet wbSheet;
        int b = 20;
        int multiplier = 0;
        int int_old_day = 1;
        String login, password;

        public MainWindow(String login, String password)
        {
            this.login = login;
            this.password = password;
            InitializeComponent();
            DateTime dateTime = DateTime.Now;
            dateTime = dateTime.AddDays(-1);
            yesterday.Content = dateTime.ToShortDateString();
            old_day.Content = "0" + int_old_day;
            inc_day.Content = "\u02C4";
            dec_day.Content = "\u02C5";

            if (flag_day)
            {
                Reinvert_flag_day();
            }
            else
            {
                Invert_flag_day();
            }
        }



        void Create_Excel_book()
        {
            Application app = new Application();
            app.Visible = true;
            if (app == null) { MessageBox.Show("Excel is not properly installed!!"); return; }
            Workbook wb = app.Workbooks.Add(Type.Missing);
            Workbook wbTemplate = app.Workbooks.Open("D:\\TemplateLERS.xls");
            wbTemplate.Sheets[1].Copy(After: wb.Worksheets[1]);
            wbTemplate.Close();
            wbSheet = (Excel.Worksheet)wb.Worksheets.get_Item(2);
        }

        void Process(DateTime date)
        {

            int[] num_heating_mains = new int[] { 21, 10, 2, 8, 3, 1, 213, 4 };
            int j = 3;
            for (int i = 0; i < num_heating_mains.Length; i++)
            {
                Heating_main heating_Main = new Heating_main(num_heating_mains[i], date, login, password);
                wbSheet.Cells[9 + b * multiplier, j] = heating_Main.T1;
                wbSheet.Cells[10 + b * multiplier, j] = heating_Main.T2;
                wbSheet.Cells[11 + b * multiplier, j] = heating_Main.D1;
                wbSheet.Cells[12 + b * multiplier, j] = heating_Main.D2;
                wbSheet.Cells[13 + b * multiplier, j] = heating_Main.dD;
                wbSheet.Cells[14 + b * multiplier, j] = heating_Main.dQ;
                wbSheet.Cells[15 + b * multiplier, j] = heating_Main.Tamur;
                j += 1;
            }

            wbSheet.Cells[7 + b * multiplier, 1] = "с " + date.ToShortDateString() + " " + date.ToShortTimeString() + "  по  " + date.ToShortDateString() + " " + (date.AddHours(24)).ToShortTimeString();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            int x1 = 7;
            int y1 = 1;
            int x2 = 24;
            int y2 = 15;
            b = 20;
            multiplier = 0;

            DateTime date_start, date_end;
            if (true)
            {
                bool flag = true;

                if (flag_day)
                {
                    DateTime dateTime = DateTime.Now;
                    date_start = dateTime.AddDays(int_old_day * (-1));
                    date_end = dateTime.AddDays(-1);
                }
                else
                {
                    if (data_start_cal.SelectedDate != null && data_end_cal.SelectedDate != null)
                    {
                        date_start = data_start_cal.SelectedDate.Value;
                        date_end = data_end_cal.SelectedDate.Value;
                    }
                    else
                    {
                        MessageBox.Show("задайте начальную и конечную дату");
                        return;
                    }
                }

                Create_Excel_book();
                wbSheet.Cells[4, 3] = " за период с  " + date_start.ToShortDateString() + " " + date_start.ToShortTimeString() + "  по  " + date_end.ToShortDateString() + " " + date_end.ToShortTimeString();

                Process(date_start);
                multiplier = 1;

                if (!System.DateTime.Equals(date_start, date_end))
                {
                    date_start = date_start.AddDays(1);
                    while (flag)
                    {

                        if (System.DateTime.Equals(date_start, date_end))
                        {
                            flag = false;
                            //break;
                        }

                        Excel.Range rng_from = wbSheet.Range[wbSheet.Cells[x1, y1], wbSheet.Cells[x2, y2]];
                        Excel.Range rng_to = wbSheet.Range[wbSheet.Cells[x1 + b * multiplier, y1], wbSheet.Cells[x2 + b * multiplier, y2]];
                        rng_from.Copy(rng_to);
                        Process(date_start);

                        multiplier++;
                        date_start = date_start.AddDays(1);

                    }
                }

            }

        }

        private void Inc_day(object sender, RoutedEventArgs e)
        {
            if (int_old_day < 30)
            {
                int_old_day++;
                old_day.Content = "0" + int_old_day;
            }
        }

        private void Dec_day(object sender, RoutedEventArgs e)
        {
            if (int_old_day > 1)
            {
                int_old_day--;
                old_day.Content = "0" + int_old_day;
            }
        }

        private void Invert_flag_day(object sender, RoutedEventArgs e)
        {
            Invert_flag_day();
            Properties.Settings.Default.dayType = false;
            Properties.Settings.Default.Save();
        }



        private void Reinvert_flag_day(object sender, RoutedEventArgs e)
        {
            Reinvert_flag_day();
            Properties.Settings.Default.dayType = true;
            Properties.Settings.Default.Save();
        }



        private void Invert_flag_day()
        {
            flag_day = false;
            data_start_cal.IsEnabled = true;
            data_end_cal.IsEnabled = true;

            inc_day.IsEnabled = false;
            dec_day.IsEnabled = false;
            old_day.IsEnabled = false;
            day_back_old.IsEnabled = false;
            yesterday.IsEnabled = false;
            
        }

        private void Reinvert_flag_day()
        {
            flag_day = true;
            data_start_cal.IsEnabled = false;
            data_end_cal.IsEnabled = false;

            inc_day.IsEnabled = true;
            dec_day.IsEnabled = true;
            old_day.IsEnabled = true;
            day_back_old.IsEnabled = true;
            yesterday.IsEnabled = true;
          
        }

    }
}

