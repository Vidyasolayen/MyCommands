using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
namespace MyCommands
{
    public partial class Configuration : Form
    {
        SqlConnection con = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=.");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        int selectedVersion;
        public Configuration()
        {
            InitializeComponent();
            DisplayData();
        }
        //Insert Data  
        private void btn_Insert_Click(object sender, EventArgs e)
        {
            if (txtVersionNumber.Text != "" && txtVersionPath.Text != "")
            {
                cmd = new SqlCommand("insert into tblVersion(versionNumber,versionPath,status, LastModifiedTimestamp) values(@versionNo,@versionPathDir,@status, @LastModifiedTimestamp)", con);
                con.Open();
                cmd.Parameters.AddWithValue("@versionNo", int.Parse(txtVersionNumber.Text));
                cmd.Parameters.AddWithValue("@versionPathDir", txtVersionPath.Text);
                cmd.Parameters.AddWithValue("@status", 1);
                cmd.Parameters.AddWithValue("@LastModifiedTimestamp", DateTime.Now);

                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Record Inserted Successfully");
                DisplayData();
                ClearData();
            }
            else
            {
                MessageBox.Show("Please Provide Details!");
            }
        }
        //Display Data in DataGridView  
        private void DisplayData()
        {
            con.Open();
            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("select versionNumber, versionPath from tblVersion", con);
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();
        }
        //Clear Data  
        private void ClearData()
        {
            txtVersionNumber.Text = "";
            txtVersionPath.Text = "";
        }
        //dataGridView1 RowHeaderMouseClick Event  
        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            txtVersionNumber.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtVersionPath.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
        }
        //Update Record  
        private void btn_Update_Click(object sender, EventArgs e)
        {
            if (txtVersionNumber.Text != "" && txtVersionPath.Text != "")
            {
                cmd = new SqlCommand("update tblVersion set versionNumber=@versionNumber,versionPath=@versionPath where versionNumber=@selectedVersion", con);
                con.Open();
                cmd.Parameters.AddWithValue("@versionNumber", int.Parse(txtVersionNumber.Text));
                cmd.Parameters.AddWithValue("@selectedVersion", selectedVersion);
                cmd.Parameters.AddWithValue("@versionPath", txtVersionPath.Text);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Record Updated Successfully");
                con.Close();
                DisplayData();
                ClearData();
            }
            else
            {
                MessageBox.Show("Please Select Record to Update");
            }
        }
        //Delete Record  
        private void btn_Delete_Click(object sender, EventArgs e)
        {
            if (txtVersionNumber.Text != "")
            {
                cmd = new SqlCommand("delete tblVersion where versionNumber=@versionNumber", con);
                con.Open();
                cmd.Parameters.AddWithValue("@versionNumber", int.Parse(txtVersionNumber.Text));
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Record Deleted Successfully!");
                DisplayData();
                ClearData();
            }
            else
            {
                MessageBox.Show("Please Select Record to Delete");
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Form1 frmMain = new Form1();
            frmMain.Show();
            this.Hide();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                selectedVersion = (int)dataGridView1.SelectedRows[0].Cells[0].Value;
                txtVersionNumber.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                txtVersionPath.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
            }
        }

        private void txtVersionPath_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtVersionNumber_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void Configuration_Load(object sender, EventArgs e)
        {

        }
    }
    }

