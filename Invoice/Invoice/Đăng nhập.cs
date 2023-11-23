using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;


using System.Net.Http;

namespace Invoice
{
    public partial class Đăng_nhập : Form
    {
        public static string UserName;
        public Đăng_nhập()
        {
            InitializeComponent();
            textBox2.PasswordChar = '*';
        }

        private async void button2_Click(object sender, EventArgs e)
        {
            var username = textBox1.Text;
            var password = textBox2.Text;
            if (username != string.Empty || password != string.Empty)
            {

                var client = new HttpClient();
                var request = new HttpRequestMessage(HttpMethod.Post, "https://admin.metalearn.vn/MobileLogin/LoginNoCheckOnline");
                var collection = new List<KeyValuePair<string, string>>();
                collection.Add(new KeyValuePair<string, string>("Username", username));
                collection.Add(new KeyValuePair<string, string>("Password", password));
                var content = new FormUrlEncodedContent(collection);
                request.Content = content;
                var response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();
                Console.WriteLine(await response.Content.ReadAsStringAsync());

                var data = await response.Content.ReadAsStringAsync();
                var datades = JsonConvert.DeserializeObject(data);

                MessageBox.Show(data);
                if (data.ToString().Contains("Đăng nhập thành công !") == true)
                {
                    UserName = username;
                    this.Hide();
                     InvoiceGateAddon registration = new InvoiceGateAddon(UserName);
                    registration.ShowDialog();
                }
                else
                {
                    Đăng_nhập.UserName = "";
                    textBox1.Clear();
                    textBox2.Clear();
                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
