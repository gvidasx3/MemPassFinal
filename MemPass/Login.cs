using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MemPass
{
    public partial class Login : Form
    {
        string connectionString = ConfigurationManager.ConnectionStrings["db_connection"].ConnectionString;
        public Login()
        { //get details from database
            InitializeComponent();
        }

        private void btnLogin_Click(object sender, EventArgs e)
        { //check if variables match for login from database and login if not give error
            if (txtEmail.Text.Length > 0 && txtLoginPass.Text.Length > 0)
            {
                SqlDataAdapter sda = new SqlDataAdapter("SELECT ID FROM Users WHERE Email='" + txtEmail.Text + "' AND Password='" + txtLoginPass.Text + "'", connectionString);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                if (dt.Rows[0][0].ToString().Length > 0)
                {
                    Dashboard dashboard = new Dashboard();
                    dashboard.UserID = dt.Rows[0][0].ToString();
                    dashboard.Show();
                    this.Hide();
                }
                else
                    MessageBox.Show("Invalid email or password !");
            }
            else
            {
                MessageBox.Show("Please enter your Email and Password !");
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            new Register().Show();
            this.Hide();
        }
    }
}
