using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace SmartCTC
{
    public partial class Remediere : Form
    {
        public Remediere()
        {
            InitializeComponent();
        }

        SqlConnection con = new SqlConnection(@"Data Source=roclut-colorful\sql2017;Initial Catalog=AplicatieCTC;Integrated Security=True");
        SqlCommand com;
        DataTable table = new DataTable();
        int indexRow;

        private void guna2GradientButton1_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new StartPage();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }

        private void guna2GradientButton2_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new SmartCTC();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }

        

        private void guna2GradientButton3_Click(object sender, EventArgs e)
        {
           
            sc.Open();
            SqlCommand cmd = sc.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "update CTC set GreseliA='" + TotalGreseliATextBox.Text + "',GreseliB='" + TotalGreseliBTextBox.Text + "',OreRemediereGreseliA='" + OreRemediereATextBox.Text + "',OreRemediereGreseliB='" + OreRemediereBTextBox.Text + "',CosturiManoperaRemediere='" + CosturiManoperaTextBox.Text + "',CosturiMaterialeRemediere='" + CosturiMaterialeTextBox.Text + "' where SerieCertificateGarantie='" + SerieCertificateTextBox.Text + "'";
            cmd.ExecuteNonQuery();
            sc.Close();

            // incarcare tabel
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

            con.Close();
            MessageBox.Show("Datele despre remediere au fost inregistrate in baza de date");
        }

        SqlConnection sc = new SqlConnection(@"Data Source=roclut-colorful\sql2017;Initial Catalog=AplicatieCTC;Integrated Security=True");
        SqlCommand cmd = new SqlCommand();
        

        private void dataGridView1_CellClick_1(object sender, DataGridViewCellEventArgs e)
        {
            indexRow = e.RowIndex;
            DataGridViewRow row = dataGridView1.Rows[indexRow];
            ClientComboBox.Text = row.Cells[1].Value.ToString();
            DenumireProiectComboBox.Text = row.Cells[2].Value.ToString();
            SerieCertificateComboBox.Text = row.Cells[3].Value.ToString();
            SerieCertificateTextBox.Text = row.Cells[4].Value.ToString();
            SerieDeclaratiiTextBox.Text = row.Cells[5].Value.ToString();
            DataTextBox.Text = row.Cells[6].Value.ToString();
            TotalGreseliATextBox.Text = row.Cells[7].Value.ToString();
            OreRemediereATextBox.Text = row.Cells[8].Value.ToString();
            TotalGreseliBTextBox.Text = row.Cells[9].Value.ToString();
            OreRemediereBTextBox.Text = row.Cells[10].Value.ToString();
            CosturiManoperaTextBox.Text = row.Cells[11].Value.ToString();
            CosturiMaterialeTextBox.Text = row.Cells[12].Value.ToString();
            ResponsabilCTCComboBox.Text = row.Cells[13].Value.ToString();

        }

        //Filtrare
        string cs = "Data Source=roclut-colorful\\sql2017;Initial Catalog=AplicatieCTC;Integrated Security=True";
        SqlConnection scon;
        SqlDataAdapter adapt;
        DataTable dt;

        private void Remediere_Load(object sender, EventArgs e)
        {
            scon = new SqlConnection(cs);
            scon.Open();
            adapt = new SqlDataAdapter("select * from CTC", scon);
            dt = new DataTable();
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            scon.Close();

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

       
        private void Remediere_Load_1(object sender, EventArgs e)
        {
            // incarcare tabel
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
            con.Close();
        }
        //export word
        public void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object toFindText, object replaceWithText)
        {

            object matchCase = true;

            object matchwholeWord = true;

            object matchwildCards = false;

            object matchSoundLike = false;

            object nmatchAllforms = false;

            object forward = true;

            object format = false;

            object matchKashida = false;

            object matchDiactitics = false;

            object matchAlefHamza = false;

            object matchControl = false;

            object read_only = false;

            object visible = true;

            object replace = -2;

            object wrap = 1;

            wordApp.Selection.Find.Execute(ref toFindText, ref matchCase,
                                            ref matchwholeWord, ref matchwildCards, ref matchSoundLike,

                                            ref nmatchAllforms, ref forward,

                                            ref wrap, ref format, ref replaceWithText,

                                                ref replace, ref matchKashida,

                                            ref matchDiactitics, ref matchAlefHamza,

                                             ref matchControl);
        }
        private void CreateWordDocument(object filename, object SaveAs)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                object missing = Missing.Value;

                Microsoft.Office.Interop.Word.Document myWordDoc = null;

                if (File.Exists((string)filename))
                {
                    object readOnly = false;

                    object isvisible = false;

                    wordApp.Visible = false;
                    myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                        ref missing, ref missing, ref missing,
                                                         ref missing, ref missing, ref missing, ref missing);
                    myWordDoc.Activate();
                    this.FindAndReplace(wordApp, "<serieDC>", SerieDeclaratiiTextBox.Text);
                    this.FindAndReplace(wordApp, "<serieCG>", SerieCertificateTextBox.Text);
                    this.FindAndReplace(wordApp, "<data>", DataTextBox.Text);
                    this.FindAndReplace(wordApp, "<serieCTC>", SerieCertificateComboBox.Text);
                    this.FindAndReplace(wordApp, "<dulap>", DenumireProiectComboBox.Text);
                    this.FindAndReplace(wordApp, "12:00:00 AM", String.Empty);


                    myWordDoc.SaveAs2(ref SaveAs, ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                                    ref missing, ref missing, ref missing);
                    wordApp.Visible = true;
                }
                MessageBox.Show("Documentul a fost generat cu succes ! Inchideti fisierul inainte de a genera altul !");
            }
            catch (Exception)
            {
                
            }
            

        }

        private void DescarcaCGRO_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            wordApp.Quit();
            CreateWordDocument(@"C:\Template\CG 2021 RO.docx", @"C:\Export\" + SerieCertificateTextBox.Text + " RO.docx");
        }

        private void DescarcaCGEN_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            wordApp.Quit();
            CreateWordDocument(@"C:\Template\CG 2021 EN.docx", @"C:\Export\" + SerieCertificateTextBox.Text + " EN.docx");
        }

        private void DescarcaDCRO_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            wordApp.Quit();
            CreateWordDocument(@"C:\Template\DC 2021 RO.docx", @"C:\Export\" + SerieDeclaratiiTextBox.Text + " RO.docx");
        }

        private void DescarcaDCEN_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            wordApp.Quit();
            CreateWordDocument(@"C:\Template\DC 2021 EN.docx", @"C:\Export\" + SerieDeclaratiiTextBox.Text + " EN.docx");
        }

       
    }
}
