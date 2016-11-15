using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using System.Data.OleDb;
using System.IO;
using System.Data;
using ExcelLibrary;
using System.Text.RegularExpressions;

namespace Total_Amount_Calculator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ExcelClass Excel;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnChoose_Click(object sender, RoutedEventArgs e)
        {
            txtPattern.Text = @"(\= \d+\.\d+)|(\= \d+)";

            //#region Test Regex

            //Regex regex = new Regex(txtPattern.Text);

            //Match match = regex.Match("Doggy ( 75 / 3 ) = 55.5 , Perl");
            //if (match.Success)
            //{

            //    MessageBox.Show(match.Value + ";");

            //}
            //else
            //    MessageBox.Show("No Match");

            //return;

            //#endregion

            if (txtPattern.Text.Trim() == "")
            {
                MessageBox.Show("Please Enter A Pattern");
                return;
            }

            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx*";

            if (fileDialog.ShowDialog() == true)
            {
                //MessageBox.Show("Hello");
                Excel = new ExcelClass(fileDialog.FileName);

                try
                {
                    using (Excel.Con)
                    {
                        Excel.Con.Open();

                        DataTable schemaTable, columnData;

                        schemaTable = new DataTable();
                        columnData = new DataTable();

                        schemaTable = Excel.Con.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

                        #region Show Names of Sheets

                        //foreach (DataRow item in schemaTable.Rows)
                        //{
                        //    MessageBox.Show(item.ItemArray[2].ToString());
                        //}

                        #endregion

                        List<string> costs = new List<string>();
                        string total = "0";
                        int percent = 0, progression = 0;
                        float monthTotal = 0;

                        foreach (DataRow row in schemaTable.Rows) // each row represents name of a sheet
                        {
                            percent += Excel.getNumberOfRows(row.ItemArray[2].ToString());
                        }

                        //MessageBox.Show(percent.ToString());                        

                        foreach (DataRow item in schemaTable.Rows) // iterating through each sheet
                        {
                            columnData = Excel.getColumnData(item.ItemArray[2].ToString(), "F1", "F2");
                            string where = "";

                            monthTotal = 0;

                            #region Calculate Total For A Day

                            foreach (DataRow dr in columnData.Rows)
                            {
                                costs = getCosts(dr.ItemArray[1].ToString());
                                total = sumOfADay(costs);
                                where = dr.ItemArray[0].ToString(); // primary key i.e. in this case date

                                if (where.Equals(""))
                                    continue;

                                try
                                {
                                    Excel.updateData(item.ItemArray[2].ToString().Replace("'", "") , "F3" , total.ToString() , where);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(item.ItemArray[2].ToString() + " | " + where + " | " + total.ToString() + "\n\n" + ex.Message);
                                }

                                barPercent.Value = (++progression * 100) / percent;
                                
                                monthTotal += float.Parse(total);                                
                            }

                            #endregion

                            try
                            {
                                Excel.updateData(item.ItemArray[2].ToString().Replace("'", ""), "F3", total.ToString() + " , Total : " + monthTotal.ToString(), where);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }

                        Excel.Con.Close();

                        MessageBox.Show("Done");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        string sumOfADay(List<string> costs)
        {
            float total = 0;

            foreach (string item in costs)
            {
                total += float.Parse(item);
            }

            return total.ToString("N1");
        }

        List<string> getCosts(string costOnADay)
        {
            List<string> costs = new List<string>();
            Regex regex = new Regex(txtPattern.Text);

            string[] split = costOnADay.Split(','); // can be changed to value from a text box later

            foreach (string item in split)
            {
                Match match = regex.Match(item);
                if (match.Success)
                {
                    costs.Add(match.Value.Replace("=", "").Replace(" ", ""));
                }
            }

            return costs;
        }
    }
}

