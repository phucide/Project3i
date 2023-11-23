using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace Invoice
{
    public partial class SelectDate : Form
    {
        InvoiceGateAddon main;
        public SelectDate(InvoiceGateAddon m)
        {
            InitializeComponent();
            main = m;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Get the selected values from ComboBoxes
            string selectedMonthFrom = comboBox1.SelectedItem as string;
            string selectedMonthTo = comboBox2.SelectedItem as string;


            // Validate that both ComboBoxes have a selected value
            if (string.IsNullOrEmpty(selectedMonthFrom))
            {
                MessageBox.Show("Please select months.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Create an object to store the selected months
            var selectedMonths = new
            {
                MonthFrom = selectedMonthFrom,
                MonthTo = selectedMonthTo
            };

            try
            {
                // Convert the object to JSON
                string json = JsonConvert.SerializeObject(selectedMonths, Formatting.Indented);

                // Specify the path where you want to save the JSON file
                string filePath = @"C:\Users\phuci\OneDrive\Desktop\Invoice-nssm-1\dist\main\SelectedMonths.json";

                // Write the JSON string to the file
                File.WriteAllText(filePath, json);

                MessageBox.Show("Selected months saved to JSON file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            main.invfrom = selectedMonthFrom;
            main.invto = selectedMonthTo;
            
            main.start_service();
            main.InitializeFileSystemWatcher();
            this.Close();
            main.Show();
            main.showWaitForm();
        }
    }
}
