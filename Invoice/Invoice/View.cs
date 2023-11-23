using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using OfficeOpenXml;
using ClosedXML.Excel;

namespace Invoice
{
    public partial class View : Form
    {
        OleDbConnection oleDbConnection = new OleDbConnection();
        private Main originalForm;
        public View(Main main)
        {
            InitializeComponent();
            originalForm = main;
            Load_Invoice();
        }

        public void Load_Invoice()
        {
            oleDbConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            @"Data Source=C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\Database.accdb;" +
            "Persist Security Info=False;";

            //oleDbConnection.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;"+@"Data Source=C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\Database.accdb;User ID=;Password=;";
            try
            {
                oleDbConnection.Open();
                MessageBox.Show("Connect Success!");

                //Remove duplicate
                string sql = "DELETE FROM Invoices_header " +
                     "WHERE id IN " +
                     "(SELECT max(id) " +
                     " FROM Invoices_header " +
                     " GROUP BY Inv_sign, Inv_num" +
                     " HAVING count(*) > 1);";

                using (OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDbConnection))
                {
                    try
                    {
                        int rowsAffected = oleDbCommand.ExecuteNonQuery();
                        Console.WriteLine($"Rows Affected: {rowsAffected}");
                    }
                    catch (OleDbException ex)
                    {
                        Console.WriteLine($"Error: {ex.Message}");
                    }
                }

                sql = "DELETE FROM Invoice_Item " +
                     "WHERE id IN " +
                     "(SELECT max(id) " +
                     " FROM Invoice_Item " +
                     " GROUP BY Invoice_code, Good_name, GTGT " +
                     " HAVING count(*) > 1);";

                using (OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDbConnection))
                {
                    try
                    {
                        int rowsAffected = oleDbCommand.ExecuteNonQuery();
                        Console.WriteLine($"Rows Affected: {rowsAffected}");
                    }
                    catch (OleDbException ex)
                    {
                        Console.WriteLine($"Error: {ex.Message}");
                    }
                }

                sql = "DELETE FROM Payment_Sum " +
                     "WHERE id IN " +
                     "(SELECT max(id) " +
                     " FROM Payment_tax " +
                     " GROUP BY Inv_code, Total_pay_num, Total_tax " +
                     " HAVING count(*) > 1);";

                using (OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDbConnection))
                {
                    try
                    {
                        int rowsAffected = oleDbCommand.ExecuteNonQuery();
                        Console.WriteLine($"Rows Affected: {rowsAffected}");
                    }
                    catch (OleDbException ex)
                    {
                        Console.WriteLine($"Error: {ex.Message}");
                    }
                }

                DataSet Invoice = new DataSet();
                sql = "select * from Invoices_header where Buy_In_Status=0;";
                OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Invoice, "Invoice");
                dataGridView1.DataSource = Invoice.Tables[0];

                DataSet Invoice_item = new DataSet();
                sql = "SELECT Good_property, Good_name, Invoice_Item.Unit, Amount, Price, Discount, Tax, GTGT, Pay, Invoice_code, Invoice_Item.No, Invoices_header.Inv_num, Invoices_header.Inv_sign, Invoices_header.Inv_date_express, Invoices_header.buyer_name, Invoices_header.buyer_tax FROM Invoices_header, Invoice_Item WHERE Invoices_header.inv_code = Invoice_Item.Invoice_code AND Invoices_header.Buy_In_Status = 0;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Invoice_item, "Invoice_item");
                dataGridView2.DataSource = Invoice_item.Tables[0];

                DataSet Payment_tax = new DataSet();
                sql = "SELECT item1,item2,item3,item4,item5,item6,item7,item8,item9,  Payment_Sum.Inv_code FROM Invoices_header, Payment_Sum WHERE Invoices_header.inv_code = Payment_Sum.Inv_code AND Invoices_header.Buy_In_Status = 0;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Payment_tax, "Payment_Sum");
                dataGridView3.DataSource = Payment_tax.Tables[0];


                Invoice = new DataSet();

                sql = "select * from Invoices_header where Buy_In_Status=1;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Invoice, "Invoice");
                dataGridView4.DataSource = Invoice.Tables[0];

                Invoice_item = new DataSet();
                sql = "SELECT Good_property, Good_name, Invoice_Item.Unit, Amount, Price, Discount, Tax, GTGT, Pay, Invoice_code, Invoice_Item.No, Invoices_header.Inv_num, Invoices_header.Inv_sign, Invoices_header.Inv_date_express, Invoices_header.buyer_name, Invoices_header.buyer_tax FROM Invoices_header, Invoice_Item WHERE Invoices_header.inv_code = Invoice_Item.Invoice_code AND Invoices_header.Buy_In_Status = 1;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Invoice_item, "Invoice_item");
                dataGridView5.DataSource = Invoice_item.Tables[0];

                Payment_tax = new DataSet();
                sql = "SELECT item1,item2,item3,item4,item5,item6,item7,item8,item9,  Payment_Sum.Inv_code FROM Invoices_header, Payment_Sum WHERE Invoices_header.inv_code = Payment_Sum.Inv_code AND Invoices_header.Buy_In_Status = 1; ";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Payment_tax, "Payment_Sum");
                dataGridView6.DataSource = Payment_tax.Tables[0];



                oleDbConnection.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        

        private void button1_Click(object sender, EventArgs e)
        {
            originalForm.Show();
            this.Hide();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            Export();


        }

        private void Export()
        {
            string existingExcelFilePath = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\Mau - To khai thue GTGT Thang 02 - 2022.xlsx";
            // Create Excel application and workbook
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet1, xlWorkSheet2, xlWorkSheet3, xlWorkSheet4;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkBook = xlexcel.Workbooks.Open(existingExcelFilePath);
            try
            {
                xlWorkSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Sheets["Bke ban ra - mau 01-1 GTGT"];
                xlWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Sheets["Bke mua vao - mau 01-2 GTGT"];
                xlWorkSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Sheets["Bang ke chi tiet hang ban ra"];
                xlWorkSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Sheets["Bang ke chi tiet hang mua vao"];
            }
            catch (Exception)
            {
                // Handle the case when the worksheet doesn't exist
                MessageBox.Show("Worksheet not found.");
                return;
            }
            CopyDataGridViewToExcel(dataGridView1, xlWorkSheet1, 9);
            CopyDataGridViewToExcel1(dataGridView4, xlWorkSheet2, 10);
            CopyDataGridViewToExcel2(dataGridView2, xlWorkSheet3, 9);
            CopyDataGridViewToExcel2(dataGridView5, xlWorkSheet4, 10);
            xlWorkBook.Save();
            xlWorkBook.Close();
            // Create the first worksheet and paste data starting from row 9


            // Create the second worksheet and paste data starting from row 9

            //xlWorkSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.Add(misValue, xlWorkSheet1);
            //xlWorkSheet2.Name = "Invoices header Buy In";
            //CopyDataGridViewToExcel(dataGridView4, xlWorkSheet2, 9);
        }

        private void CopyDataGridViewToExcel(DataGridView dataGridView, Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, int startRow)
        {
            // Copy column headers
            //for (int j = 0; j < dataGridView.Columns.Count; j++)
            //{
            //    xlWorkSheet.Cells[startRow, j + 1] = dataGridView.Columns[j].HeaderText;
            //}

            // Copy data from DataGridView to Excel
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                DataGridViewRow row = dataGridView.Rows[i];



                //for (int j = 0; j < row.Cells.Count; j++)
                //{
                //    xlWorkSheet.Cells[startRow + i + 1, j + 1] = row.Cells[j].Value?.ToString();
                //}
                xlWorkSheet.Cells[startRow + i, 1] = row.Cells[2].Value?.ToString();
                xlWorkSheet.Cells[startRow+i,2] = row.Cells[6].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 3] =  row.Cells[5].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 4] = row.Cells[12].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 5] = row.Cells[14].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 6] = "";
                xlWorkSheet.Cells[startRow + i, 7] = "";
                xlWorkSheet.Cells[startRow + i, 8] = "";
                xlWorkSheet.Cells[startRow + i, 9] = "";
            }
        }

        private void CopyDataGridViewToExcel1(DataGridView dataGridView, Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, int startRow)
        {
            // Copy column headers
            //for (int j = 0; j < dataGridView.Columns.Count; j++)
            //{
            //    xlWorkSheet.Cells[startRow, j + 1] = dataGridView.Columns[j].HeaderText;
            //}

            // Copy data from DataGridView to Excel
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                DataGridViewRow row = dataGridView.Rows[i];



                //for (int j = 0; j < row.Cells.Count; j++)
                //{
                //    xlWorkSheet.Cells[startRow + i + 1, j + 1] = row.Cells[j].Value?.ToString();
                //}
                xlWorkSheet.Cells[startRow + i, 1] = row.Cells[2].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 2] = row.Cells[6].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 3] = row.Cells[5].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 4] = row.Cells[7].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 5] = row.Cells[8].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 6] = "";
                xlWorkSheet.Cells[startRow + i, 7] = "";
                xlWorkSheet.Cells[startRow + i, 8] = "";
                xlWorkSheet.Cells[startRow + i, 9] = "";
                xlWorkSheet.Cells[startRow + i, 10] = "";
            }
        }

        private void CopyDataGridViewToExcel2(DataGridView dataGridView, Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, int startRow)
        {
            // Copy column headers
            //for (int j = 0; j < dataGridView.Columns.Count; j++)
            //{
            //    xlWorkSheet.Cells[startRow, j + 1] = dataGridView.Columns[j].HeaderText;
            //}

            // Copy data from DataGridView to Excel
            for (int i = 0; i < dataGridView.Rows.Count; i++)
            {
                DataGridViewRow row = dataGridView.Rows[i];



                //for (int j = 0; j < row.Cells.Count; j++)
                //{
                //    xlWorkSheet.Cells[startRow + i + 1, j + 1] = row.Cells[j].Value?.ToString();
                //}
                xlWorkSheet.Cells[startRow + i, 1] = row.Cells[11].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 2] = row.Cells[12].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 3] = row.Cells[9].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 4] = row.Cells[13].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 5] = row.Cells[14].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 6] = row.Cells[15].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 7] = row.Cells[1].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 8] = row.Cells[3].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 9] = row.Cells[4].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 10] = "";
                xlWorkSheet.Cells[startRow + i, 11] = "";
                xlWorkSheet.Cells[startRow + i, 12] = "";

            }
        }


    }
}
