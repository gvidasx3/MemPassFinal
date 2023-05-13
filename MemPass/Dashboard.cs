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
using Excel = Microsoft.Office.Interop.Excel;

namespace MemPass
{
    public partial class Dashboard : Form
    { //load data from database
        string connectionString = ConfigurationManager.ConnectionStrings["db_connection"].ConnectionString;
        public string UserID { get; set; }
        public Dashboard()
        {
            InitializeComponent();
        }

        private void Dashboard_Load(object sender, EventArgs e)
        {
            
            loadData();          
        }

        private void button4_Click(object sender, EventArgs e)
        { //Update password button
            if (textBox7.Text.Length > 0)
            {
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand cmd = new SqlCommand("UPDATE Users SET Password = '"+ textBox7.Text + "' Where ID="+ Convert.ToInt32(UserID), con);
                cmd.ExecuteNonQuery();
                con.Close();

                MessageBox.Show("Password updated Successfully !");
                clearData();
                loadData();

            }
            else
            { //If password same
                MessageBox.Show("Please enter new password !");
            }
            
        }

        private void btnLogout_Click(object sender, EventArgs e)
        { //Logout button sends you to fresh start of Login page
            new Login().Show();
            this.Hide();
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        { // Textboxes turn input to string
            var index = e.RowIndex;
            textBox1.Text = dataGridView1.Rows[index].Cells[0].Value.ToString();
            textBox2.Text = dataGridView1.Rows[index].Cells[1].Value.ToString();
            textBox3.Text = dataGridView1.Rows[index].Cells[2].Value.ToString();
            textBox4.Text = dataGridView1.Rows[index].Cells[3].Value.ToString();
            textBox5.Text = dataGridView1.Rows[index].Cells[4].Value.ToString();
            textBox6.Text = dataGridView1.Rows[index].Cells[5].Value.ToString();

            if (dataGridView1.Rows[index].Cells[0].Value.ToString().Length > 0)
            { // When a password from vault is selected with mouse show delete and update button
                button1.Enabled = false;

                button2.Enabled = true;
                button3.Enabled = true;
            }

            else
            { // No password selected only show insert button
                button1.Enabled = true;

                button2.Enabled = false;
                button3.Enabled = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        { //Delete password button
            try
            {
                SqlConnection con = new SqlConnection(connectionString);
                con.Open();
                SqlCommand cmd = new SqlCommand("Delete from Passwords Where ID = " + Convert.ToInt32(textBox1.Text), con);
                cmd.ExecuteNonQuery();
                con.Close();

                MessageBox.Show("Password deleted Successfully !");
                clearData();
                loadData();

                button1.Enabled = true;

                button2.Enabled = false;
                button3.Enabled = false;

            }
            catch (Exception)
            {
                MessageBox.Show("Error !");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        { //Update password in vault + database
            if (textBox2.Text.Length > 0 && textBox3.Text.Length > 0 && textBox4.Text.Length > 0 && textBox5.Text.Length > 0 && textBox6.Text.Length > 0)
            {
                try
                {
                    SqlConnection con = new SqlConnection(connectionString);
                    con.Open();
                    SqlCommand cmd = new SqlCommand("Update Passwords set Tag = '"+ textBox2.Text + "', Password = '" + textBox3.Text + "', SQ1 = '" + textBox4.Text + "', SQ2 = '" + textBox5.Text + "', SQ3 = '" + textBox6.Text + "' Where ID = " + Convert.ToInt32(textBox1.Text), con);
                    cmd.ExecuteNonQuery();
                    con.Close();

                    MessageBox.Show("Password updated Successfully !");
                    clearData();
                    loadData();

                    button1.Enabled = true;

                    button2.Enabled = false;
                    button3.Enabled = false;

                }
                catch (Exception)
                {
                    MessageBox.Show("Error !");
                }
            }
            else
            {
                MessageBox.Show("Please enter values in all fields !");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.Length > 0 && textBox3.Text.Length > 0 && textBox4.Text.Length > 0 && textBox5.Text.Length > 0 && textBox6.Text.Length > 0)
            { //if all fields are filled and insert is pressed, enter password in vault + database
                //try
                //{
                    SqlConnection con = new SqlConnection(connectionString);
                    con.Open();
                    SqlCommand cmd = new SqlCommand("Insert into Passwords(UserID,Tag,Password,SQ1,SQ2,SQ3) Values(" + Convert.ToInt32(UserID) + ", '" + textBox2.Text + "', '" + textBox3.Text + "', '" + textBox4.Text + "', '" + textBox5.Text + "', '" + textBox6.Text + "')", con);
                    cmd.ExecuteNonQuery();
                    con.Close();

                    MessageBox.Show("Password inserted Successfully !");
                    clearData();
                    loadData();

                    button1.Enabled = true;

                    button2.Enabled = false;
                    button3.Enabled = false;

                //}
                //catch (Exception)
                //{
                //    MessageBox.Show("Error !");
                //}
            }
            else
            {
                MessageBox.Show("Please enter values in all fields !");
            }

        }

        public void clearData()
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            //checkBox1.Enabled = false;
            //checkBox2.Enabled = false;
        }

        public void loadData()
        { //Loads Tag, password etc from database
            SqlDataAdapter da = new SqlDataAdapter("SELECT ID, Tag, Password, SQ1, SQ2, SQ3 FROM Passwords Where UserID='" + this.UserID + "';", connectionString);
            DataSet ds = new DataSet();
            da.Fill(ds, "Passwords");
            dataGridView1.DataSource = ds.Tables["Passwords"].DefaultView;
        }

        public char GetRandomSpecialCharacter()
        { //special char contains
            string chars = "$%#@!*?;:^&~";
            Random rand = new Random();
            int num = rand.Next(0, chars.Length);
            return chars[num];
        }

        public char GetRandomAlphaNumeric()
        { //alpha numerics contains
            string chars = "abcdefghijklmnopqrstuvwxyz1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            Random rand = new Random();
            int num = rand.Next(0, chars.Length);
            return chars[num];
        }

        private void button5_Click(object sender, EventArgs e)
        { //Combine question field answers to generate password by scrambling the words
            if (textBox4.Text.Length > 0 && textBox5.Text.Length > 0 && textBox6.Text.Length > 0)
            {
                string password = "";
                string[] security_questions = { textBox4.Text, textBox5.Text, textBox6.Text };
                Random random = new Random();
                security_questions = security_questions.OrderBy(x => random.Next()).ToArray();

                foreach (var item in security_questions)
                {
                    if (checkBox1.Checked)
                    { //If special characters checked add to password
                        password += GetRandomSpecialCharacter();
                    }

                    password += item;

                    if (checkBox2.Checked)
                    { //If Alpha characters checked add to password
                        password += GetRandomAlphaNumeric();
                    }
                }

                textBox3.Text = password;
            }
            else
            {
                MessageBox.Show("Please answer security questions !");
            }


        }

        private void btnExportPasswords_Click(object sender, EventArgs e)
        { //To export user's vault as an excel file to F drive
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            int i = 0;
            int j = 0;

            for (i = 0; i <= dataGridView1.RowCount - 1; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dataGridView1[j, i];
                    xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
                }
            }
            //Creates file name as "Passwords" with values from vault
            xlWorkBook.SaveAs(@"F:\Passwords.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Passwords Excel file created at F:");
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }
    }
}
