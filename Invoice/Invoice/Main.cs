using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.ServiceProcess;
using System.Data.OleDb;

namespace Invoice
{
    public partial class Main : Form

    {
        public static Main instance;
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

        //private bool stopBackgroundWorker = false;

        public Main(string userName)
        {
            username = userName;
            invfrom = "";
            invto = "";
            InitializeComponent();
            GetwindowServices();
            checkService();
            waitForm = new WaitFormFunction();
            instance = this;

        }

        public void checkService()
        {
            if (service_name.Equals("Invoice-nssm"))
            {
                button1.Text = "Remove";
                ServiceController single_service = new ServiceController(service_name);
                if (single_service.Status.Equals(ServiceControllerStatus.Stopped) || single_service.Status.Equals(ServiceControllerStatus.StopPending))
                {
                    button3.Enabled = false;
                    button4.Enabled = true;
                }
                else
                {
                    button3.Enabled = true;
                    button4.Enabled = false;
                }
            }
            else
            {
                button1.Text = "Install";
                button3.Enabled = false;
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

        public void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
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



        private void button4_Click(object sender, EventArgs e)
        {
            //stopBackgroundWorker = false;
            //SelectDate s = new SelectDate(this);
            //s.Show();
            //this.Hide();


            startTime = DateTime.Now;
            
            
            
            textBox1.Clear();
            this.button4.Enabled = false;
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
                    Invoke(new Action(() => textBox1.Text = fileContent));

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
            this.button4.Enabled = true;
            this.button3.Enabled = false;

            // Check if the background worker is busy before starting it
            
            this.backgroundWorker1.RunWorkerAsync(service_name);
            
            textBox1.AppendText("Service stop!!");

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

            
            View view = new View(this);
            view.Show();
            this.Hide();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string nssmInstallCommand = "";
            string nssm_log_file = "";
            string service_Name = "Invoice-nssm";
            if (service_name.Equals("Invoice-nssm"))
            {
                nssmInstallCommand = $"nssm remove " + service_Name;
                button1.Text = "Install";

            }
            else
            {
                string executablePath = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\dist\main\main.exe";  // Replace with the path to your executable

                // Prepare the nssm command to install the service
                nssmInstallCommand = $"nssm install " + service_Name + " " + executablePath;
                string log_file = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\service.log";
                nssm_log_file = $"nssm set Invoice-nssm AppStdout " +log_file;
                button1.Text = "Remove";
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

        
    }
}
