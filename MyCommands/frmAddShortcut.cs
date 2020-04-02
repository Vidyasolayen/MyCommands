using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyCommands
{
    public partial class frmAddShortcut : Form
    {
        SqlConnection con = new SqlConnection("Integrated Security=SSPI;Persist Security Info=False;User ID=wbpoc;Initial Catalog=DFCommands;Data Source=.");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        int selectedShortcutId;

        public frmAddShortcut()
        {
            InitializeComponent();
            DisplayData();
        }

        public void showCodeOnScreen(string code)
        {
            txtCodeList.AppendText($"{Environment.NewLine}{code}");
        }

    
        //Insert Data  
        private void btn_Insert_Click(object sender, EventArgs e)
        {
            if (txtShortcutName.Text != "" && txtCode.Text != "")
            {
                cmd = new SqlCommand("insert into tblShortcuts(ShortName,ShortcutCode) values(@name, @code)", con);
                con.Open();
                cmd.Parameters.AddWithValue("@code", int.Parse(txtCode.Text));
                cmd.Parameters.AddWithValue("@name", txtShortcutName.Text);

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
            adapt = new SqlDataAdapter("select ShortcutId, ShortName, ShortcutCode from tblShortcuts", con);
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();
        }
        //Clear Data  
        private void ClearData()
        {
            txtShortcutName.Text = "";
            txtCode.Text = "";
            txtShortcutId.Text = "";
        }
        //dataGridView1 RowHeaderMouseClick Event  
        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            txtShortcutName.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtCode.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
        }
        //Update Record  
        private void btn_Update_Click(object sender, EventArgs e)
        {
            if (txtShortcutName.Text != "" && txtCode.Text != "")
            {
                cmd = new SqlCommand("update tblShortcuts set ShortName=@name,ShortcutCode=@code where ShortcutId=@id", con);
                con.Open();
                cmd.Parameters.AddWithValue("@id", selectedShortcutId);
                cmd.Parameters.AddWithValue("@name", txtShortcutName.Text);
                cmd.Parameters.AddWithValue("@code", txtCode.Text);
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
            if (txtShortcutName.Text != "")
            {
                cmd = new SqlCommand("delete tblShortcuts where ShortcutId=@shortcutId", con);
                con.Open();
                cmd.Parameters.AddWithValue("@shortcutId", int.Parse(txtShortcutId.Text));
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                selectedShortcutId = (int)dataGridView1.SelectedRows[0].Cells[0].Value;
                txtShortcutName.Text = dataGridView1.SelectedRows[0].Cells[0].Value.ToString();
                txtCode.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Form1 frmMain = new Form1();
            frmMain.Show();
            this.Hide();
        }

        private void frmAddShortcut_Load(object sender, EventArgs e)
        {

        }
    }
}
