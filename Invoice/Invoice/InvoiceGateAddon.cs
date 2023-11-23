using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Threading;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Data.OleDb;
using System.Xml;

namespace Invoice
{
    public partial class InvoiceGateAddon : Form
    {
        public static InvoiceGateAddon instance;
        string filePath = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\service.log";
        private FileSystemWatcher fileWatcher;
        public string service_name = "";
        WaitFormFunction waitForm = new WaitFormFunction();
        string username = "";
        OleDbConnection oleDbConnection = new OleDbConnection();
        DateTime startTime;
        DateTime endTime;
        public string invfrom;
        public string invto;
        public bool flag_Check_Account;
        public InvoiceGateAddon(string userName)
        {
            InitializeComponent();
            flag_Check_Account = false;
            username = userName;
            invfrom = "";
            invto = "";
            GetwindowServices();
            checkService();
            waitForm = new WaitFormFunction();
            instance = this;
            label9.Text = userName;
            textBox4.PasswordChar = '*';
            Load_Invoice();
        }

        public void checkService()
        {
            if (flag_Check_Account==true)
            {
                button6.BackColor = Color.Green;
                button4.Enabled = true;
                if (service_name.Equals("Invoice-nssm"))
                {
                    button4.Text = "Remove";
                    ServiceController single_service = new ServiceController(service_name);
                    if (single_service.Status.Equals(ServiceControllerStatus.Stopped) || single_service.Status.Equals(ServiceControllerStatus.StopPending))
                    {
                        button3.Enabled = false;
                        button2.Enabled = true;
                    }
                    else
                    {
                        button3.Enabled = true;
                        button2.Enabled = false;
                    }
                }
                else
                {
                    button4.Text = "Install";
                    button3.Enabled = false;
                    button2.Enabled = false;
                }
            }
            else
            {
                
                button3.Enabled = false;
                button2.Enabled = false;
                button4.Enabled = false;
            }
            

        }

        private void GetwindowServices()
        {
            //throw new NotImplementedException();
            ServiceController[] service;
            service = ServiceController.GetServices();
            for (int i = 0; i < service.Length; i++)
            {
                string serviceName = service[i].ServiceName;
                if (serviceName.Equals("Invoice-nssm"))
                    service_name = serviceName;
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var single_service = new ServiceController((string)e.Argument);

            if (single_service.Status.Equals(ServiceControllerStatus.Stopped) || single_service.Status.Equals(ServiceControllerStatus.StopPending))
            {
                single_service.Start();

            }
            else
            {
                single_service.Stop();

            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // Perform cleanup and open the View form only once
            if (e.Error == null)
            {
                //textBox1.AppendText("Process complete!!! Please press the stop button");

                // Open the View form if not already open

            }
            else
            {
                MessageBox.Show(e.Error.ToString(), "Cannot start service!!!");
            }
        }

        public void start_service()
        {
            this.backgroundWorker1.RunWorkerAsync(service_name);

        }

        public void showWaitForm()
        {
            waitForm.Show(this);
        }

        public void stopWaitForm()
        {
            waitForm.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //stopBackgroundWorker = false;
            SelectDate s = new SelectDate(this);
            s.Show();
            this.Hide();

            startTime = DateTime.Now;

            textBox1.Clear();
            this.button2.Enabled = false;
            this.button3.Enabled = true;
        }

        public void InitializeFileSystemWatcher()
        {
            fileWatcher = new FileSystemWatcher(Path.GetDirectoryName(filePath));
            fileWatcher.Filter = Path.GetFileName(filePath);

            // Subscribe to the events
            fileWatcher.Changed += OnFileChanged;
            fileWatcher.EnableRaisingEvents = true;
        }

        private void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            // Handle the file change event
            UpdateTextBoxContent();

        }
        private void UpdateTextBoxContent()
        {
            const int maxAttempts = 10;
            const int retryInterval = 1000; // 1 second

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    // Attempt to read the content of the file
                    string fileContent = File.ReadAllText(filePath);

                    // Invoke to update the TextBox on the UI thread
                    Invoke(new Action(() => textBox6.Text = fileContent));

                    stopWaitForm();

                    // Check if the desired content is present in the file


                    // Break out of the loop if reading is successful
                    break;
                }
                catch (IOException ex)
                {
                    // Log or display the exception message
                    Console.WriteLine($"Error reading the file (Attempt {attempt}/{maxAttempts}): {ex.Message}");

                    // Wait before attempting to read the file again
                    Thread.Sleep(retryInterval);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            endTime = DateTime.Now;
            Console.WriteLine("Current Time: " + endTime);
            Console.WriteLine("Username: " + username);
            this.button2.Enabled = true;
            this.button3.Enabled = false;

            // Check if the background worker is busy before starting it

            this.backgroundWorker1.RunWorkerAsync(service_name);

            textBox6.AppendText("Service stop!!");

            const int maxAttempts = 10;
            const int retryInterval = 1000; // 1 second

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    File.WriteAllText(filePath, string.Empty);
                    // Break out of the loop if reading is successful
                    break;
                }
                catch (IOException ex)
                {
                    // Wait before attempting to read the file again
                    Thread.Sleep(retryInterval);
                }
            }

            oleDbConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source=C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\Database.accdb;" +
                "Persist Security Info=False;";

            try
            {
                string sql = "INSERT INTO Log_table(Username, starttime, endtime, invfromtime, invtotime) VALUES (?, ?, ? , ? , ?);";

                using (OleDbCommand oleDbCommand = new OleDbCommand(sql, oleDbConnection))
                {
                    oleDbCommand.Parameters.AddWithValue("@Username", username);
                    oleDbCommand.Parameters.AddWithValue("@StartTime", startTime);
                    oleDbCommand.Parameters.AddWithValue("@EndTime", endTime);
                    oleDbCommand.Parameters.AddWithValue("@invfromtime", invfrom);
                    oleDbCommand.Parameters.AddWithValue("@invtotime", invto);

                    try
                    {
                        oleDbConnection.Open();
                        int rowsAffected = oleDbCommand.ExecuteNonQuery();
                        Console.WriteLine($"Rows Affected: {rowsAffected}");
                    }
                    catch (OleDbException ex)
                    {
                        Console.WriteLine($"Error: {ex.Message}");
                    }
                    finally
                    {
                        oleDbConnection.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //View view = new View(this);
            //view.Show();
            //this.Hide();
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
                     " GROUP BY Inv_sign, Inv_num, inv_code" +
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
                     " FROM Payment_Sum " +
                     " GROUP BY Inv_code, item8 " +
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

                sql = "DELETE FROM Tax " +
                     "WHERE id IN " +
                     "(SELECT max(id) " +
                     " FROM Tax " +
                     " GROUP BY Inv_code, Total_no_tax, Total_tax " +
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

                sql = "DELETE FROM Fee " +
                     "WHERE id IN " +
                     "(SELECT max(id) " +
                     " FROM Fee " +
                     " GROUP BY Inv_code, Fee_name, Fee " +
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
                dataGridView5.DataSource = Invoice.Tables[0];

                DataSet Invoice_tax = new DataSet();
                sql = "SELECT Tax,Total_no_tax,Total_tax,Tax.Inv_code FROM Invoices_header, Tax WHERE Invoices_header.inv_code = Tax.Inv_code AND Invoices_header.Buy_In_Status = 0;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Invoice_tax, "Invoice_Tax");
                dataGridView6.DataSource = Invoice_tax.Tables[0];

                DataSet Invoice_item = new DataSet();
                sql = "SELECT Good_property, Good_name, Invoice_Item.Unit, Amount, Price, Discount, Tax, GTGT, Pay, Invoice_code, Invoice_Item.No, Invoices_header.Inv_num, Invoices_header.Inv_sign, Invoices_header.Inv_date_express, Invoices_header.buyer_name, Invoices_header.buyer_tax FROM Invoices_header, Invoice_Item WHERE Invoices_header.inv_code = Invoice_Item.Invoice_code AND Invoices_header.Buy_In_Status = 0;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Invoice_item, "Invoice_item");
                dataGridView7.DataSource = Invoice_item.Tables[0];

                DataSet Payment_tax = new DataSet();
                sql = "SELECT item1,item2,item3,item4,item5,item6,item7,item8,item9,  Payment_Sum.Inv_code FROM Invoices_header, Payment_Sum WHERE Invoices_header.inv_code = Payment_Sum.Inv_code AND Invoices_header.Buy_In_Status = 0;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Payment_tax, "Payment_Sum");
                dataGridView8.DataSource = Payment_tax.Tables[0];


                Invoice = new DataSet();

                sql = "select * from Invoices_header where Buy_In_Status=1;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Invoice, "Invoice");
                dataGridView1.DataSource = Invoice.Tables[0];

                Invoice_tax = new DataSet();
                sql = "SELECT Tax,Total_no_tax,Total_tax,Tax.Inv_code FROM Invoices_header, Tax WHERE Invoices_header.inv_code = Tax.Inv_code AND Invoices_header.Buy_In_Status = 1;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Invoice_tax, "Invoice_Tax");
                dataGridView2.DataSource = Invoice_tax.Tables[0];

                Invoice_item = new DataSet();
                sql = "SELECT Good_property, Good_name, Invoice_Item.Unit, Amount, Price, Discount, Tax, GTGT, Pay, Invoice_code, Invoice_Item.No, Invoices_header.Inv_num, Invoices_header.Inv_sign, Invoices_header.Inv_date_express, Invoices_header.buyer_name, Invoices_header.buyer_tax FROM Invoices_header, Invoice_Item WHERE Invoices_header.inv_code = Invoice_Item.Invoice_code AND Invoices_header.Buy_In_Status = 1;";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Invoice_item, "Invoice_item");
                dataGridView3.DataSource = Invoice_item.Tables[0];

                Payment_tax = new DataSet();
                sql = "SELECT item1,item2,item3,item4,item5,item6,item7,item8,item9,  Payment_Sum.Inv_code FROM Invoices_header, Payment_Sum WHERE Invoices_header.inv_code = Payment_Sum.Inv_code AND Invoices_header.Buy_In_Status = 1; ";
                oleDbDataAdapter = new OleDbDataAdapter(sql, oleDbConnection);
                oleDbDataAdapter.Fill(Payment_tax, "Payment_Sum");
                dataGridView4.DataSource = Payment_tax.Tables[0];



                oleDbConnection.Dispose();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string nssmInstallCommand = "";
            string nssm_log_file = "";
            string service_Name = "Invoice-nssm";
            if (service_name.Equals("Invoice-nssm"))
            {
                nssmInstallCommand = $"nssm remove " + service_Name;
                button4.Text = "Install";

            }
            else
            {
                string executablePath = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\dist\main\main.exe";  // Replace with the path to your executable

                // Prepare the nssm command to install the service
                nssmInstallCommand = $"nssm install " + service_Name + " " + executablePath;
                string log_file = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\service.log";
                nssm_log_file = $"nssm set Invoice-nssm AppStdout " + log_file;
                button4.Text = "Remove";
            }
            using (Process process = new Process())
            {
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;

                process.Start();

                // Send the nssm command to the Command Prompt
                process.StandardInput.WriteLine(nssmInstallCommand);
                process.StandardInput.WriteLine(nssm_log_file);
                process.StandardInput.WriteLine("exit");

                // Read the output (if needed)
                string output = process.StandardOutput.ReadToEnd();

                process.WaitForExit();

                // Display the output or handle it as needed
                MessageBox.Show("Service installation complete.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            service_name = "";
            GetwindowServices();
            checkService();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Load_Invoice();
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
            CopyDataGridViewToExcel(dataGridView5, xlWorkSheet1, 9);
            CopyDataGridViewToExcel1(dataGridView1, xlWorkSheet2, 10);
            CopyDataGridViewToExcel2(dataGridView7, xlWorkSheet3, 9);
            CopyDataGridViewToExcel2(dataGridView3, xlWorkSheet4, 10);
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
                xlWorkSheet.Cells[startRow + i, 2] = row.Cells[6].Value?.ToString();
                xlWorkSheet.Cells[startRow + i, 3] = row.Cells[5].Value?.ToString();
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

        private void InvoiceGateAddon_Load(object sender, EventArgs e)

        {
            timer1.Start();
            label8.Text = DateTime.Now.ToLongTimeString();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            label8.Text = DateTime.Now.ToLongTimeString();
            timer1.Start();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            
            if (textBox4.Text.Length<8 || textBox4.Text.Length >15)
            {
                MessageBox.Show("The password must be in range[8,15]", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (textBox3.Text != "" && textBox4.Text != "")
            {
                UpdateXmlFile();
                CallPythonToCheckAccount();
            }
            else
            {
                MessageBox.Show("Please fill Mã số thuế and Mật khẩu", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void CallPythonToCheckAccount()
        {
            string executablePath = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\dist\main\Check_Account.exe";

            // Create a ProcessStartInfo instance
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = executablePath,
                WorkingDirectory = System.IO.Path.GetDirectoryName(executablePath), // Set working directory to executable's directory
                UseShellExecute = false, // Ensure that the executable is not run using the shell
                RedirectStandardOutput = true, // Redirect standard output if needed
                CreateNoWindow = false // Do not create a window for the process
            };

            // Start the process
            using (Process process = new Process { StartInfo = startInfo })
            {
                process.Start();

                // Optionally, read the output of the process
                string output = process.StandardOutput.ReadToEnd();
                Console.WriteLine(output);

                process.WaitForExit();
            }

            string filePath = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\dist\main\Check_Account.txt";
            string status = "";
            try
            {
                // Read the entire file as a single string
                status = File.ReadAllText(filePath);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            if(status == "Sucess")
            {
                flag_Check_Account = true;
                MessageBox.Show("Login sucess", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                checkService();
            }
            else
            {
                flag_Check_Account = false;
                button6.BackColor = Color.Red;
                MessageBox.Show("Login Fail!!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                checkService();
            }
        }

        private void UpdateXmlFile()
        {
            
            string xmlFilePath = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\dist\main\Invoice_Setting.xml";
            try
            {
                // Load the XML document
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlFilePath);

                // Find the username element
                XmlNode usernameNode = xmlDoc.SelectSingleNode("/invoice/variable/uername");
                XmlNode passNode = xmlDoc.SelectSingleNode("/invoice/variable/password");

                // Update the text content
                if (usernameNode != null)
                {
                    usernameNode.InnerText = textBox3.Text; // Replace with the desired username
                }
                else
                {
                    MessageBox.Show("Username element not found in XML.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (passNode != null)
                {
                    passNode.InnerText = textBox4.Text; // Replace with the desired username
                }
                else
                {
                    MessageBox.Show("Username element not found in XML.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Save the changes to the XML file
                xmlDoc.Save(xmlFilePath);

                MessageBox.Show("XML file updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating XML file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            
                
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var item = dataGridView1.Rows[e.RowIndex].Cells[6].Value?.ToString();
                oleDbConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            @"Data Source=C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\Database.accdb;" +
            "Persist Security Info=False;";

                if (!string.IsNullOrEmpty(item))
                {
                    using (OleDbConnection oleDbConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;" +
                 @"Data Source=C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\Database.accdb;" +
                 "Persist Security Info=False;"))
                    {
                        oleDbConnection.Open();

                        // 1. Invoice Tax
                        DataSet Invoice_tax = new DataSet();
                        string taxQuery = "SELECT * FROM Tax WHERE Inv_code = @ItemCode";
                        using (OleDbCommand taxCommand = new OleDbCommand(taxQuery, oleDbConnection))
                        {
                            taxCommand.Parameters.AddWithValue("@ItemCode", item);
                            OleDbDataAdapter taxDataAdapter = new OleDbDataAdapter(taxCommand);
                            taxDataAdapter.Fill(Invoice_tax, "Invoice_Tax");
                            dataGridView2.DataSource = Invoice_tax.Tables[0];
                        }

                        // 2. Invoice Item
                        DataSet Invoice_item = new DataSet();
                        string itemQuery = "SELECT * FROM Invoice_item WHERE Invoice_code = @ItemCode";
                        using (OleDbCommand itemCommand = new OleDbCommand(itemQuery, oleDbConnection))
                        {
                            itemCommand.Parameters.AddWithValue("@ItemCode", item);
                            OleDbDataAdapter itemDataAdapter = new OleDbDataAdapter(itemCommand);
                            itemDataAdapter.Fill(Invoice_item, "Invoice_item");
                            dataGridView3.DataSource = Invoice_item.Tables[0];
                        }

                        // 3. Payment Tax
                        DataSet Payment_tax = new DataSet();
                        string paymentQuery = "SELECT * FROM Payment_Sum WHERE Inv_code = @ItemCode";
                        using (OleDbCommand paymentCommand = new OleDbCommand(paymentQuery, oleDbConnection))
                        {
                            paymentCommand.Parameters.AddWithValue("@ItemCode", item);
                            OleDbDataAdapter paymentDataAdapter = new OleDbDataAdapter(paymentCommand);
                            paymentDataAdapter.Fill(Payment_tax, "Payment_Sum");
                            dataGridView4.DataSource = Payment_tax.Tables[0];
                        }
                    }
                }
            }
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                var item = dataGridView5.Rows[e.RowIndex].Cells[6].Value?.ToString();
                oleDbConnection.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
            @"Data Source=C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\Database.accdb;" +
            "Persist Security Info=False;";

                if (!string.IsNullOrEmpty(item))
                {
                    using (OleDbConnection oleDbConnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;" +
                 @"Data Source=C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\Database.accdb;" +
                 "Persist Security Info=False;"))
                    {
                        oleDbConnection.Open();

                        // 1. Invoice Tax
                        DataSet Invoice_tax = new DataSet();
                        string taxQuery = "SELECT * FROM Tax WHERE Inv_code = @ItemCode";
                        using (OleDbCommand taxCommand = new OleDbCommand(taxQuery, oleDbConnection))
                        {
                            taxCommand.Parameters.AddWithValue("@ItemCode", item);
                            OleDbDataAdapter taxDataAdapter = new OleDbDataAdapter(taxCommand);
                            taxDataAdapter.Fill(Invoice_tax, "Invoice_Tax");
                            dataGridView6.DataSource = Invoice_tax.Tables[0];
                        }

                        // 2. Invoice Item
                        DataSet Invoice_item = new DataSet();
                        string itemQuery = "SELECT * FROM Invoice_item WHERE Invoice_code = @ItemCode";
                        using (OleDbCommand itemCommand = new OleDbCommand(itemQuery, oleDbConnection))
                        {
                            itemCommand.Parameters.AddWithValue("@ItemCode", item);
                            OleDbDataAdapter itemDataAdapter = new OleDbDataAdapter(itemCommand);
                            itemDataAdapter.Fill(Invoice_item, "Invoice_item");
                            dataGridView7.DataSource = Invoice_item.Tables[0];
                        }

                        // 3. Payment Tax
                        DataSet Payment_tax = new DataSet();
                        string paymentQuery = "SELECT * FROM Payment_Sum WHERE Inv_code = @ItemCode";
                        using (OleDbCommand paymentCommand = new OleDbCommand(paymentQuery, oleDbConnection))
                        {
                            paymentCommand.Parameters.AddWithValue("@ItemCode", item);
                            OleDbDataAdapter paymentDataAdapter = new OleDbDataAdapter(paymentCommand);
                            paymentDataAdapter.Fill(Payment_tax, "Payment_Sum");
                            dataGridView8.DataSource = Payment_tax.Tables[0];
                        }
                    }
                }
            }
        }
    }
}
