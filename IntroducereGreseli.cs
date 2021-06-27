using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;
using System.Data.Odbc;

namespace SmartCTC
{
    public partial class SmartCTC : Form
    {

        public SmartCTC()
        {
            InitializeComponent();
            FillCombo();
            DenumireProiectComboBox.Enabled = false;
            guna2GradientButton5.Enabled = false;
            AdaugaResponsabil.Enabled = false;
            guna2GradientButton3.Enabled = false;
        }

        SqlConnection sc = new SqlConnection(@"Data Source=roclut-colorful\sql2017;Initial Catalog=AplicatieCTC;Integrated Security=True");
        DataTable table = new DataTable("table");
        DataTable table1 = new DataTable("table1");
        int index;
        void Serii()
        {
            string constr = @"Data Source=roclut-colorful\sql2017;Initial Catalog=AplicatieCTC;Integrated Security=True";

            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT max(cast(right(SerieCertificateGarantie,len(SerieCertificateGarantie)-8) as int))+1 as num from CTC where len(SerieCertificateGarantie)>8"))
                {
                    DateTime now = DateTime.Today;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = con;
                    con.Open();
                    using (SqlDataReader sdr = cmd.ExecuteReader())
                    {
                        string a = string.Empty;
                        sdr.Read();
                        a = sdr["num"].ToString();
                        int b = Int32.Parse(a);
                        a = b.ToString();
                        SerieCertificateTextBox.Text = ("CG " + now.ToString("yyyy") + "-" + a ).ToString(); ;
                    }
                    con.Close();
                }
            }

            using (SqlConnection con = new SqlConnection(constr))
            {
                using (SqlCommand cmd = new SqlCommand("SELECT max(cast(right(SerieDeclaratiiDeConformitate,len(SerieDeclaratiiDeConformitate)-8) as int))+1 as num from CTC where len(SerieDeclaratiiDeConformitate)>8"))
                {
                    DateTime now = DateTime.Today;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = con;
                    con.Open();
                    using (SqlDataReader sdr = cmd.ExecuteReader())
                    {
                        string a = string.Empty;
                        sdr.Read();
                        a = sdr["num"].ToString();
                        int b = Int32.Parse(a);
                        a = b.ToString();
                        SerieDeclaratiiTextBox.Text = ("DC " + now.ToString("yyyy") + "-" + a ).ToString();
                    }
                    con.Close();
                }
            }

        }
        void FillCombo()

        {

            //umplere seriectc
            string timesheetconstring = "Dsn=timesheet;uid=a.bogdan";
            OdbcConnection timesheetconodbc = new OdbcConnection(timesheetconstring);
            using (OdbcConnection timesheetconodb = new OdbcConnection(timesheetconstring))
            {
                try
                {
                    string query = "select * from projects ORDER BY code DESC";
                    OdbcDataAdapter da = new OdbcDataAdapter(query, timesheetconodb);
                    timesheetconodb.Open();
                    DataSet ds = new DataSet();
                    da.Fill(ds, "projects");
                    SerieCertificateComboBox.DisplayMember = "code";
                    SerieCertificateComboBox.ValueMember = "code";
                    SerieCertificateComboBox.DataSource = ds.Tables["projects"];
                    SerieCertificateComboBox.SelectedIndex = -1;
                    SerieCertificateComboBox.Text = "Selectati";
                }
                catch (Exception ex)
                {
                    // write exception info to log or anything else
                    MessageBox.Show("Conexiunea cu odbc - timesheet!");
                }
            }

            timesheetconodbc.Close();

        }

        private void Add_DataGridView(object sender, EventArgs e)
        {
            table.Columns.Add("", typeof(int));
            table1.Columns.Add("", typeof(int));
        }


        private void SmartCTC_Load(object sender, EventArgs e)
        {
            table.Columns.Add("Client", Type.GetType("System.String"));
            table.Columns.Add("Denumire proiect", Type.GetType("System.String"));
            table.Columns.Add("Serie certificate CTC", Type.GetType("System.String"));
            table.Columns.Add("Serie certificate garantie", Type.GetType("System.String"));
            table.Columns.Add("Serie declaratii de conformitate", Type.GetType("System.String"));
            table.Columns.Add("Data", Type.GetType("System.DateTime"));
            table.Columns.Add("Greseli A", Type.GetType("System.Int32"));
            table.Columns.Add("Ore remediere greseli A", Type.GetType("System.Decimal"));
            table.Columns.Add("Greseli B", Type.GetType("System.Int32"));
            table.Columns.Add("Ore remediere greseli B", Type.GetType("System.Decimal"));
            table.Columns.Add("Costuri manopera remediere", Type.GetType("System.Int32"));
            table.Columns.Add("Costuri materiale remediere", Type.GetType("System.Int32"));
            table.Columns.Add("Responsabil CTC", Type.GetType("System.String"));
            dataGridView1.DataSource = table;
            table1.Columns.Add("Responsabil", Type.GetType("System.String"));
            table1.Columns.Add("Proiect", Type.GetType("System.String"));
            table1.Columns.Add("Greseli A", Type.GetType("System.String"));
            table1.Columns.Add("Greseli B", Type.GetType("System.String"));
            dataGridView2.DataSource = table1;

        }



        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            index = e.RowIndex;
            DataGridViewRow row = dataGridView1.Rows[index];
            ClientComboBox.Text = row.Cells[0].Value.ToString();
            DenumireProiectComboBox.Text = row.Cells[1].Value.ToString();
            SerieCertificateComboBox.Text = row.Cells[2].Value.ToString();
            SerieCertificateTextBox.Text = row.Cells[3].Value.ToString();
            SerieDeclaratiiTextBox.Text = row.Cells[4].Value.ToString();
            TotalGreseliATextBox.Text = row.Cells[5].Value.ToString();
            TotalGreseliBTextBox.Text = row.Cells[7].Value.ToString();
            CosturiManoperaTextBox.Text = row.Cells[9].Value.ToString();
            CosturiMaterialeTextBox.Text = row.Cells[10].Value.ToString();
            ResponsabilCTCComboBox.Text = row.Cells[11].Value.ToString();
        }


        private void guna2GradientButton3_Click(object sender, EventArgs e)
        {

            table.Rows.Add(ClientComboBox.Text, DenumireProiectComboBox.Text, SerieCertificateComboBox.Text, SerieCertificateTextBox.Text, SerieDeclaratiiTextBox.Text, DataDateTime.Value.Date, TotalGreseliATextBox.Text, OreRemediereATextBox.Text, TotalGreseliBTextBox.Text, OreRemediereBTextBox.Text, CosturiManoperaTextBox.Text, CosturiMaterialeTextBox.Text, ResponsabilCTCComboBox.Text);
            guna2GradientButton5.Enabled = true;
            guna2GradientButton3.Enabled = false;
            AdaugaResponsabil.Enabled = false;




        }

        private void guna2GradientButton4_Click(object sender, EventArgs e)
        {
            index = dataGridView1.CurrentCell.RowIndex;
            dataGridView1.Rows.RemoveAt(index);
            MessageBox.Show("Randul a fost sters");
        }

        private void guna2GradientButton5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                SqlCommand cmd = new SqlCommand(@"INSERT INTO CTC (Client,DenumireProiect,SerieCertificateCTC,SerieCertificateGarantie,SerieDeclaratiiDeConformitate,Data,GreseliA,OreRemediereGreseliA,GreseliB,OreRemediereGreseliB,CosturiManoperaRemediere,CosturiMaterialeRemediere,ResponsabilCTC)VALUES('" + dataGridView1.Rows[i].Cells[0].Value + "','" + dataGridView1.Rows[i].Cells[1].Value + "','" + dataGridView1.Rows[i].Cells[2].Value + "','" + dataGridView1.Rows[i].Cells[3].Value + "','" + dataGridView1.Rows[i].Cells[4].Value + "','" + dataGridView1.Rows[i].Cells[5].Value + "','" + dataGridView1.Rows[i].Cells[6].Value + "','" + dataGridView1.Rows[i].Cells[7].Value + "','" + dataGridView1.Rows[i].Cells[8].Value + "','" + dataGridView1.Rows[i].Cells[9].Value + "','" + dataGridView1.Rows[i].Cells[10].Value + "','" + dataGridView1.Rows[i].Cells[11].Value + "','" + dataGridView1.Rows[i].Cells[12].Value + "')", sc);
                sc.Open();
                cmd.ExecuteNonQuery();
                sc.Close();
            }
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                SqlCommand cmd = new SqlCommand(@"INSERT INTO ResponsabilExecutie (Responsabil,Proiect,GreseliA,GreseliB)VALUES('" + dataGridView2.Rows[i].Cells[0].Value + "','" + dataGridView2.Rows[i].Cells[1].Value + "','" + dataGridView2.Rows[i].Cells[2].Value + "','" + dataGridView2.Rows[i].Cells[3].Value + "')", sc);
                sc.Open();
                cmd.ExecuteNonQuery();
                sc.Close();
            }
            MessageBox.Show("Greselile au fost inregistrate in baza de date");
            //golire gridview & CG DC
            do
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    try
                    {
                        dataGridView1.Rows.Remove(row);
                    }
                    catch (Exception) { }
                }
            } while (dataGridView1.Rows.Count > 0);

            do
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    try
                    {
                        dataGridView2.Rows.Remove(row);
                    }
                    catch (Exception) { }
                }
            } while (dataGridView2.Rows.Count > 0);

            guna2GradientButton5.Enabled = false;
            SerieCertificateTextBox.Text = string.Empty;
            SerieDeclaratiiTextBox.Text = string.Empty;
        }

        private void guna2GradientButton1_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new StartPage();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }

        private void guna2GradientButton6_Click(object sender, EventArgs e)
        {
            this.Hide();
            var form2 = new Remediere();
            form2.Closed += (s, args) => this.Close();
            form2.Show();
        }
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
            try {
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
                    this.FindAndReplace(wordApp, "<data>", DataDateTime.Value.Date.ToString("d"));
                    this.FindAndReplace(wordApp, "<serieCTC>", SerieCertificateComboBox.Text);
                    this.FindAndReplace(wordApp, "<dulap>", DenumireProiectComboBox.Text);



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

        

        private void DescarcaCG_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            wordApp.Quit();
            CreateWordDocument(@"C:\Template\CG 2021 RO.docx", @"C:\Export\" + SerieCertificateTextBox.Text + " RO.docx");

        }

        private void DescarcaDC_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            wordApp.Quit();
            CreateWordDocument(@"C:\Template\DC 2021 RO.docx", @"C:\Export\" + SerieDeclaratiiTextBox.Text + " RO.docx");
        }

        private void DescarcaCGEN_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            wordApp.Quit();
            CreateWordDocument(@"C:\Template\CG 2021 EN.docx", @"C:\Export\" + SerieCertificateTextBox.Text + " EN.docx");
        }

        private void DescarcaDCEN_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            object missing = Missing.Value;
            wordApp.Quit();
            CreateWordDocument(@"C:\Template\DC 2021 EN.docx", @"C:\Export\" + SerieDeclaratiiTextBox.Text + " EN.docx");
        }

        private void SerieCertificateComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SerieCertificateTextBox.Text = string.Empty;
            SerieDeclaratiiTextBox.Text = string.Empty;


            //umplere client
            string clienticonstring = "Dsn=clienti;database=clientzi;option=0;port=3306;server=192.168.1.26;uid=a.bogdan";
            OdbcConnection clienticonodbc = new OdbcConnection(clienticonstring);
            using (OdbcConnection clienticonodb = new OdbcConnection(clienticonstring))
            {
                try
                {
                    //
                    string query = "select client from cerere_oferte where nr_intrare = left('" + SerieCertificateComboBox.Text + "',7)";
                    OdbcDataAdapter da = new OdbcDataAdapter(query, clienticonodb);
                    clienticonodbc.Open();
                    DataSet ds = new DataSet();
                    da.Fill(ds, "cerere_oferte");
                    ClientComboBox.DisplayMember = "client";
                    ClientComboBox.ValueMember = "client";
                    ClientComboBox.DataSource = ds.Tables["cerere_oferte"];
                }
                catch (Exception ex)
                {
                    // write exception info to log or anything else
                    MessageBox.Show("Conexiunea cu odbc - clienti!");
                }
            }

            string str = "Select name from projects Where code ='" + SerieCertificateComboBox.SelectedValue + "'";
            using (OdbcConnection con = new OdbcConnection(@"Dsn=timesheet;uid=a.bogdan"))
            {
                using (OdbcCommand cmd = new OdbcCommand(str, con))
                {
                    cmd.Parameters.AddWithValue("name", DenumireProiectComboBox.Text);
                    using (OdbcDataAdapter adp = new OdbcDataAdapter(cmd))
                    {
                        DataTable dtItem = new DataTable();
                        adp.Fill(dtItem);
                        DenumireProiectComboBox.DataSource = dtItem;
                        DenumireProiectComboBox.DisplayMember = "name";
                        DenumireProiectComboBox.ValueMember = "name";
                    }
                }
            }

            //umplere CG si DC
            string CGQuery = "select distinct SerieCertificateGarantie from CTC where SerieCertificateCTC = '" + SerieCertificateComboBox.SelectedValue + "'";
            using (SqlConnection con = new SqlConnection("Data Source=roclut-colorful\\sql2017;Initial Catalog=AplicatieCTC;Integrated Security=True"))
            {
                using (SqlCommand cmd = new SqlCommand(CGQuery, con))
                {
                    cmd.Parameters.AddWithValue("SerieCertificateGarantie", SerieCertificateTextBox.Text);
                    using (SqlDataAdapter adp = new SqlDataAdapter(cmd))
                    {
                        DataTable dtItem = new DataTable();
                        adp.Fill(dtItem);
                        SerieCertificateTextBox.DataSource = dtItem;
                        SerieCertificateTextBox.DisplayMember = "SerieCertificateGarantie";
                        SerieCertificateTextBox.ValueMember = "SerieCertificateGarantie";
                    }
                }
            }
            string DCQuery = "select distinct SerieDeclaratiiDeConformitate from CTC where SerieCertificateCTC = '" + SerieCertificateComboBox.SelectedValue + "'";
            using (SqlConnection con = new SqlConnection("Data Source=roclut-colorful\\sql2017;Initial Catalog=AplicatieCTC;Integrated Security=True"))
            {
                using (SqlCommand cmd = new SqlCommand(DCQuery, con))
                {
                    cmd.Parameters.AddWithValue("SerieDeclaratiiDeConformitate", SerieDeclaratiiTextBox.Text);
                    using (SqlDataAdapter adp = new SqlDataAdapter(cmd))
                    {
                        DataTable dtItem = new DataTable();
                        adp.Fill(dtItem);
                        SerieDeclaratiiTextBox.DataSource = dtItem;
                        SerieDeclaratiiTextBox.DisplayMember = "SerieDeclaratiiDeConformitate";
                        SerieDeclaratiiTextBox.ValueMember = "SerieDeclaratiiDeConformitate";
                    }
                }
            }

            DenumireProiectComboBox.Enabled = true;
            AdaugaResponsabil.Enabled = true;
            SerieCertificateTextBox.Enabled = true;
            SerieDeclaratiiTextBox.Enabled = true;

        }
        private void DenumireProiectComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string str1 = "Select distinct t1.EmployeeId,t1.ProjectCode,t2.regnb,t2.firstname from inputdata as t1 inner join employees as t2 on t1.EmployeeId=t2.regnb Where t1.ProjectCode ='" + SerieCertificateComboBox.SelectedValue + "'";
            using (OdbcConnection con = new OdbcConnection(@"Dsn=timesheet;uid=a.bogdan"))
            {
                using (OdbcCommand cmd = new OdbcCommand(str1, con))
                {
                    cmd.Parameters.AddWithValue("firstname", ResponsabilComboBox.Text);
                    using (OdbcDataAdapter adp = new OdbcDataAdapter(cmd))
                    {
                        DataTable dtItem = new DataTable();
                        adp.Fill(dtItem);
                        ResponsabilComboBox.DataSource = dtItem;
                        ResponsabilComboBox.DisplayMember = "firstname";
                        ResponsabilComboBox.ValueMember = "firstname";
                    }
                }
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            index = e.RowIndex;
            DataGridViewRow row = dataGridView2.Rows[index];
            ResponsabilComboBox.Text = row.Cells[0].Value.ToString();
            SerieCertificateComboBox.Text = row.Cells[1].Value.ToString();
            GreseliAComboBox.Text = row.Cells[2].Value.ToString();
            GreseliBComboBox.Text = row.Cells[3].Value.ToString();
        }

        private void AdaugaResponsabil_Click(object sender, EventArgs e)
        {
            table1.Rows.Add(ResponsabilComboBox.Text, SerieCertificateComboBox.Text, GreseliAComboBox.Text, GreseliBComboBox.Text);
            double sumA = 0;
            for (int i = 0; i < dataGridView2.Rows.Count; ++i)
            {
                sumA += Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
            }
            TotalGreseliATextBox.Text = sumA.ToString();
            double sumB = 0;
            for (int i = 0; i < dataGridView2.Rows.Count; ++i)
            {
                sumB += Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);
            }
            TotalGreseliBTextBox.Text = sumB.ToString();
            guna2GradientButton3.Enabled = true;
        }

        private void GenereazaCGDCButton_Click(object sender, EventArgs e)
        {
            Serii();

        }
    }
}
