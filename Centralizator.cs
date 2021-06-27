using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Windows.Data;

namespace SmartCTC
{
    public partial class Centralizator : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=roclut-colorful\sql2017;Initial Catalog=AplicatieCTC;Integrated Security=True");
        SqlCommand com;

        public Centralizator()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            
        }
        

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new StartPage();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }

        

        

        private void guna2GradientButton2_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new StartPage();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }

        string cs = "Data Source=roclut-colorful\\sql2017;Initial Catalog=AplicatieCTC;Integrated Security=True";
        SqlConnection scon;
        SqlDataAdapter adapt;
        DataTable dt;

        private void Centralizator_Load(object sender, EventArgs e)
        {

            scon = new SqlConnection(cs);
            scon.Open();
            adapt = new SqlDataAdapter("select * from CTC", scon);
            dt = new DataTable();
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            scon.Close();

            // calculare total si afisare

            com = new SqlCommand();

            com.Connection = con;

            com.CommandText = "select * from CTC";

            con.Open();



            SqlDataReader reader = com.ExecuteReader();

            if (reader.HasRows)

            {

                DataTable dt = new DataTable();

                dt.Load(reader);

                dataGridView1.DataSource = dt;

            }

            //total greseli
            TotalATextBox.Text = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[7].FormattedValue.ToString() != string.Empty select Convert.ToInt32(row.Cells[7].FormattedValue)).Sum().ToString();
            TotalBTextBox.Text = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[9].FormattedValue.ToString() != string.Empty select Convert.ToInt32(row.Cells[9].FormattedValue)).Sum().ToString();

            decimal a, b;
            a = decimal.Parse(TotalATextBox.Text);
            b = decimal.Parse(TotalBTextBox.Text);
            TotalGreseliTextBox.Text = (a + b).ToString();
            //total ore
            string adresaA, adresaB;
            double doubA, doubB, s;

            adresaA = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[8].FormattedValue.ToString() != string.Empty select Convert.ToDecimal(row.Cells[8].FormattedValue)).Sum().ToString();
            adresaB = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[10].FormattedValue.ToString() != string.Empty select Convert.ToDecimal(row.Cells[10].FormattedValue)).Sum().ToString();

            doubA = double.Parse(adresaA);
            doubB = double.Parse(adresaB);

            var timeX = TimeSpan.FromMinutes(doubA);
            var timeY = TimeSpan.FromMinutes(doubB);

            TotalOreRemediereATextBox.Text = timeX.ToString();
            TotalOreRemediereBTextBox.Text = timeY.ToString();

            s = doubA + doubB;

            var timeSpan = TimeSpan.FromMinutes(s);

            TotalOreRemediereTextBox.Text = timeSpan.ToString();

            //total costuri

            CostManoperaRemediereTextBox.Text = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[11].FormattedValue.ToString() != string.Empty select Convert.ToDecimal(row.Cells[11].FormattedValue)).Sum().ToString();
            CostMaterialeRemediereTextBox.Text = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[12].FormattedValue.ToString() != string.Empty select Convert.ToDecimal(row.Cells[12].FormattedValue)).Sum().ToString();

            decimal coManRem, coMatRem;

            coManRem = decimal.Parse(CostManoperaRemediereTextBox.Text);
            coMatRem = decimal.Parse(CostMaterialeRemediereTextBox.Text);

            TotalCosturiRemediereTextBox.Text = (coManRem + coMatRem).ToString();
            con.Close();
        }

        private void ClientFilter_TextChanged(object sender, EventArgs e)
        {
            scon = new SqlConnection(cs);
            scon.Open();
            adapt = new SqlDataAdapter("select * from CTC where Client like '" + ClientFilter.Text + "%'", scon);
            dt = new DataTable();
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            scon.Close();
        }

        private void SerieCTCFilter_TextChanged(object sender, EventArgs e)
        {
            scon = new SqlConnection(cs);
            scon.Open();
            adapt = new SqlDataAdapter("select * from CTC where SerieCertificateCTC like '" + SerieCTCFilter.Text + "%'", scon);
            dt = new DataTable();
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            scon.Close();

            
        }

        private void DataFilter_TextChanged(object sender, EventArgs e)
        {
            scon = new SqlConnection(cs);
            scon.Open();
            adapt = new SqlDataAdapter("select * from CTC where Data like '" + DataFilter.Text + "%'", scon);
            dt = new DataTable();
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            scon.Close();
        }

        private void DulapFilter_TextChanged(object sender, EventArgs e)
        {
            scon = new SqlConnection(cs);
            scon.Open();
            adapt = new SqlDataAdapter("select * from CTC where DenumireProiect like '" + DulapFilter.Text + "%'", scon);
            dt = new DataTable();
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            scon.Close();
        }

        private void SerieCTCFilter_Enter(object sender, EventArgs e)
        {
            com = new SqlCommand();

            com.Connection = con;

            com.CommandText = "select * from CTC";

            con.Open();

            SqlDataReader reader = com.ExecuteReader();

            if (reader.HasRows)

            {

                DataTable dt = new DataTable();

                dt.Load(reader);

                dataGridView1.DataSource = dt;

            }

            //total greseli
            TotalATextBox.Text = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[7].FormattedValue.ToString() != string.Empty select Convert.ToInt32(row.Cells[7].FormattedValue)).Sum().ToString();
            TotalBTextBox.Text = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[9].FormattedValue.ToString() != string.Empty select Convert.ToInt32(row.Cells[9].FormattedValue)).Sum().ToString();

            decimal a, b;
            a = decimal.Parse(TotalATextBox.Text);
            b = decimal.Parse(TotalBTextBox.Text);
            TotalGreseliTextBox.Text = (a + b).ToString();
            //total ore
            string adresaA, adresaB;
            double doubA, doubB, s;

            adresaA = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[8].FormattedValue.ToString() != string.Empty select Convert.ToDecimal(row.Cells[8].FormattedValue)).Sum().ToString();
            adresaB = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[10].FormattedValue.ToString() != string.Empty select Convert.ToDecimal(row.Cells[10].FormattedValue)).Sum().ToString();

            doubA = double.Parse(adresaA);
            doubB = double.Parse(adresaB);

            var timeX = TimeSpan.FromMinutes(doubA);
            var timeY = TimeSpan.FromMinutes(doubB);

            TotalOreRemediereATextBox.Text = timeX.ToString();
            TotalOreRemediereBTextBox.Text = timeY.ToString();

            s = doubA + doubB;

            var timeSpan = TimeSpan.FromMinutes(s);

            TotalOreRemediereTextBox.Text = timeSpan.ToString();

            //total costuri

            CostManoperaRemediereTextBox.Text = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[11].FormattedValue.ToString() != string.Empty select Convert.ToDecimal(row.Cells[11].FormattedValue)).Sum().ToString();
            CostMaterialeRemediereTextBox.Text = (from DataGridViewRow row in dataGridView1.Rows where row.Cells[12].FormattedValue.ToString() != string.Empty select Convert.ToDecimal(row.Cells[12].FormattedValue)).Sum().ToString();

            decimal coManRem, coMatRem;

            coManRem = decimal.Parse(CostManoperaRemediereTextBox.Text);
            coMatRem = decimal.Parse(CostMaterialeRemediereTextBox.Text);

            TotalCosturiRemediereTextBox.Text = (coManRem + coMatRem).ToString();
            con.Close();
        }
    }
}
