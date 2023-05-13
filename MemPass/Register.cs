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
    public partial class Register : Form
    { 
        string connectionString = ConfigurationManager.ConnectionStrings["db_connection"].ConnectionString;
        public Register()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            new Login().Show();
            this.Hide();
        }

        private void btnRegister_Click(object sender, EventArgs e)
        { //create user in database
            if (txtEmail.Text.Length > 0 && txtLoginPass.Text.Length > 0)
            {
                try
                {
                    SqlConnection conn = new SqlConnection(connectionString);
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("Insert into Users values('" + txtEmail.Text + "','" + txtLoginPass.Text + "');", conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    MessageBox.Show("User Registered Successfully !");

                    new Login().Show();
                    this.Hide();
                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error !");
                }
            }
            else
            {
                MessageBox.Show("Please enter your Email and Password !");
            }
        }
    }
}
