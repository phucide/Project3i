using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Net.Http;

namespace Invoice
{
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
            textBox2.PasswordChar = '*';
        }

        public static string UserName;

        private async void button1_Click(object sender, EventArgs e)
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
                    Main registration = new Main(UserName);
                    registration.ShowDialog();
                }
                else
                {
                    Login.UserName = "";
                    textBox1.Clear();
                    textBox2.Clear();
                }

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        
    }
}
