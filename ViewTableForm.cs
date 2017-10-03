using ayoti.ServiceReference1;//ayoti
//using ayoti.ServiceReference2;//zagaa
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Data.SqlClient;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;

namespace ayoti
{
    public partial class ViewTableForm : Form
    {
        class SaleP
        {
            public string Customer { get; set; }
            public string SalesPerson { get; set; }
        }
        class SP
        {
            public string Code { get; set; }
            public string SalesPerson { get; set; }
        }
        private string routecode;
        List<SP> sps = new List<SP>()  {
             new SP() { Code = "AY016", SalesPerson = "James Ogwalla" },
             new SP() { Code = "AY018", SalesPerson = "Isaac Wandera" },
             new SP() { Code = "AY019", SalesPerson = "Ezekiel Owino" },
             new SP() { Code = "AY020", SalesPerson = "Harrison Wandera" },
             new SP() { Code = "AY021", SalesPerson = "Wilson Mukuna" },
             new SP() { Code = "AY022", SalesPerson = "Tobias Otieno" },
             new SP() { Code = "AY024", SalesPerson = "Bob Ndong'a" },
             new SP() { Code = "AY028", SalesPerson = "Michael Oduor" },
             new SP() { Code = "AY029", SalesPerson = "Lameck Oketch" },
             new SP() { Code = "AY035", SalesPerson = "Gordon Otieno" },
        };
        List<SP> routes = new List<SP>()  {
             new SP() { Code = "R001", SalesPerson = "Mashinani/Rural" },
             new SP() { Code = "R737", SalesPerson = "Staff" },
             new SP() { Code = "R782", SalesPerson = "Lake Key accounts" },
             new SP() { Code = "R066", SalesPerson = "Kondele" },
             new SP() { Code = "R067", SalesPerson = "Town" },
             new SP() { Code = "R071", SalesPerson = "Khayega" },
             new SP() { Code = "R072", SalesPerson = "Khemaio" },
             new SP() { Code = "R073", SalesPerson = "Luanda" },
             new SP() { Code = "R074", SalesPerson = "Mbale" },
             new SP() { Code = "R075", SalesPerson = "Serem" },
             new SP() { Code = "R075", SalesPerson = "Kapsabet" },
             new SP() { Code = "R077", SalesPerson = "Muhoroni" },
             new SP() { Code = "R078", SalesPerson = "Gem Alego" },
             new SP() { Code = "R079", SalesPerson = "Uyoma" },
             new SP() { Code = "R080", SalesPerson = "Yimbo" },
        };
        private SqlConnection con;
        string DBCon = ConfigurationManager.ConnectionStrings["ayoti.Properties.Settings.AyotiConnectionString"].ConnectionString;
        List<Sale> saleslst = new List<Sale>();
        List<Sales_Invoice> salesInvoicelst = new List<Sales_Invoice>();
        IEnumerable<SaleP> customerlst = null;

        public ViewTableForm()
        {
            InitializeComponent();
        }

        private void salesBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.salesBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.ayotiDataSet);

        }

        private void ViewTableForm_Load(object sender, EventArgs e)
        {
            getLocations();
            this.salesTableAdapter.Fill(this.ayotiDataSet.Sales);
            saleslst = this.ayotiDataSet.Tables[0].AsEnumerable().Select(d => new Sale
            {
                Id = d.Field<int>("Id"),
                Product = d.Field<string>("Product"),
                Category = d.Field<string>("Category"),
                Quantity = d.Field<double>("Quantity"),
                Customer = d.Field<string>("Customer"),
                SalesPerson = d.Field<string>("SalesPerson"),
                Uploaded = d.Field<bool>("Uploaded")
            }).Where(d => d.Uploaded == false).ToList();
            customerlst = saleslst.GroupBy(x => new { x.Customer, x.SalesPerson })
                .Select(grp => new SaleP()
               {
                   Customer = grp.First().Customer,
                   SalesPerson = grp.First().SalesPerson
               }).ToList();
            loadRoutes();
            cmbRoute.Text = "Mashinani/Rural";
            
        }

        void loadRoutes()
        {
            foreach (SP route in routes)
            {
                cmbRoute.Items.Add(route.SalesPerson);
            }
        }
        private void btnUpload_Click(object sender, EventArgs e)
        {
            btnUpload.Enabled = false;
            btnBBack.Enabled = false;
            label1.Text = "Started Upload...";
            //NAV nav = new NAV(new Uri("https://zagaasportsfoundationorg.financials.dynamics.com:7048/MS/OData/Company('AYO')"));
            //nav.Credentials = new NetworkCredential("TIMOTHY", "U8Blq49KHpLTYHJe97dDdWBN71GROlm1xjbV78qsSTU=");
            NAV nav = new NAV(new Uri("https://ayotigroup.financials.dynamics.com:7048/MS/OData/Company('AYOTI LIVE')"));//AYOTI LIVE
            nav.Credentials = new NetworkCredential("ADMIN", "sdG6zobPh0dw8vUNcAGJ4tkMuCZhX1IO8w5K+pAu7ng=");
            routecode = routes.Where(r => r.SalesPerson.Equals(cmbRoute.Text)).FirstOrDefault().Code;
            // PrintCustomersCalledCust(nav);
            CreateInvoice(nav);
            //MockCreateInvoice(nav);
        }

        public void getLocations()
        {
            DataTable dt = new DataTable();

            using (var con = new SqlConnection(DBCon))
            {
                con.Open();

                try
                {
                    SqlCommand cmd = new SqlCommand("SELECT location FROM locations", con);
                    dt.Load(cmd.ExecuteReader());
                }
                catch (SqlException e)
                {
                    Console.WriteLine(e.ToString());
                    return;
                }
            }
            
            cmbLocation.DataSource = dt;
            cmbLocation.ValueMember = dt.Columns[0].ColumnName; ;
            cmbLocation.DisplayMember = dt.Columns[0].ColumnName;
        }

        void MockCreateInvoice(NAV nav)
        {
            Sales_Invoice invoice = null;
            Sales_InvoiceSalesLines line = null;
               
                invoice = new Sales_Invoice()
                {
                    Document_Type = "Invoice",
                    Sell_to_Customer_No = "KE0001340",//"10000",
                    Salesperson_Code = "AY016",
                    Posting_Date = dateTimePicker.Value,
                    //Currency_Code = "KES",
                    //live
                  //  Shortcut_Dimension_1_Code = "R737",
                    //other
                    Shortcut_Dimension_2_Code = "AY016",
                    Location_Code = "KAJ295 Y",
                };
                nav.AddObject("Sales_Invoice", invoice);

            line = new Sales_InvoiceSalesLines()
            {
                Description = "IT00004",
                //"1025",
                Quantity = (Decimal)2
            };
                    nav.AddRelatedObject(invoice, "Sales_InvoiceSalesLines", line);
                
                try
                {
                    DataServiceResponse response = nav.SaveChanges();
                    if (response != null)
                    {
                        int uploaded = setUploaded("KE0001340", "KE0001340");
                        if (uploaded < 0)
                        {
                            string ex = string.Format("Did not update any records with customer: {0} and salesperson: {1}", "KE0001340", "KE0001340");
                            throw new System.InvalidOperationException(ex);
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    var confirmResult = MessageBox.Show("Do you want to exit the application ??",
                                     "Confirm Exit",
                                     MessageBoxButtons.YesNo);
                    if (confirmResult == DialogResult.Yes)
                    {
                        // If 'Yes', do something here.
                        Application.Exit();
                        return;
                    }
                    else
                    {
                        // If 'No', do something here.
                    }
                }
                finally
                {

                }
            
            btnUpload.Enabled = true;
            btnBBack.Enabled = true;
            label1.Text = "Completed Upload";
        }

        void CreateInvoice(NAV nav)
        {
            Sales_Invoice invoice = null;
            Sales_InvoiceSalesLines line = null;
            foreach (var c in customerlst)
            {
                Sale currentsale = null;
                List<Sale> salesline = saleslst.Where(x => x.Customer == c.Customer).ToList();
                invoice = new Sales_Invoice()
                {
                    Document_Type = "Invoice",
                    Sell_to_Customer_No = c.Customer.Trim(),
                    Salesperson_Code = getSalesPerson(c.SalesPerson.Trim()),
                    Posting_Date = dateTimePicker.Value,
                    Shipment_Date = dateTimePicker.Value,
                    Due_Date = dateTimePicker.Value,    
                    //Currency_Code = "KES",
                    //live
                    Shortcut_Dimension_1_Code = routecode,
                    //other
                    Shortcut_Dimension_2_Code = getSalesPerson(c.SalesPerson.Trim()),
                    Location_Code = cmbLocation.Text
                };
                nav.AddObject("Sales_Invoice", invoice);
                foreach (var saleline in salesline)
                {
                    currentsale = saleline;
                    line = new Sales_InvoiceSalesLines()
                    {
                        Description = saleline.Product,
                        Quantity = (Decimal)saleline.Quantity                        
                    };
                    nav.AddRelatedObject(invoice, "Sales_InvoiceSalesLines", line);
                }
                try
                {
                    DataServiceResponse response = nav.SaveChanges();
                    if (response != null)
                    {
                        int uploaded = setUploaded(c.Customer, c.SalesPerson.Trim());
                        if(uploaded< 0)
                        {
                            string ex = string.Format("Did not update any records with customer: {0} and salesperson: {1}", c.Customer, c.SalesPerson);
                            throw new System.InvalidOperationException(ex);
                        }
                    }
                    
                }
                catch ( Exception ex)
                {
                    if(currentsale != null)
                    {
                        String errorstr = String.Format("Customer :{0} with Product {1} Quantity {2} .Error: ", c.Customer, currentsale.Product, currentsale.Quantity);
                        Logger.LogThisLine(errorstr + ex.Message);
                    }                    
                    MessageBox.Show(ex.ToString());
                    var confirmResult = MessageBox.Show("Do you want to exit the application ??",
                                     "Confirm Exit",
                                     MessageBoxButtons.YesNo);
                    if (confirmResult == DialogResult.Yes)
                    {
                        // If 'Yes', do something here.
                        Application.Exit();
                        return;
                    }
                    else
                    {
                        // If 'No', do something here.
                    }
                }
                finally
                {

                }                
            }
            btnUpload.Enabled = true;
            btnBBack.Enabled = true;
            label1.Text = "Completed Upload";
        }

        int setUploaded(string customer,string salesperson)
        {
            int rowsUpdated = 0;
            using (con = new SqlConnection(DBCon))
            {
                con.Open();
                string sql = "UPDATE Sales SET Uploaded=1 WHERE Customer =@customer AND SalesPerson=@salesperson";
                
                using (SqlCommand command = new SqlCommand(sql, con))
                {
                    command.Parameters.AddWithValue("@customer", customer);
                    command.Parameters.AddWithValue("@salesperson", salesperson);
                    rowsUpdated = command.ExecuteNonQuery();
                }
            }
            return rowsUpdated;
        }

        string getSalesPerson(string salesp)
        {
            return sps.Where(x => x.SalesPerson == salesp).Select(x => x.Code).FirstOrDefault();
        }

        private void btnFetch_Click(object sender, EventArgs e)
        {
            //// TODO: This line of code loads data into the 'ayotiDataSet.Sales' table. You can move, or remove it, as needed.
            //this.salesTableAdapter.Fill(this.ayotiDataSet.Sales);
            //saleslst = this.ayotiDataSet.Tables[0].AsEnumerable().Select(d => new Sale
            //{
            //    Id = d.Field<int>("Id"),
            //    Product = d.Field<string>("Product"),
            //    Category = d.Field<string>("Category"),
            //    Quantity = d.Field<double>("Quantity"),
            //    Customer = d.Field<string>("Customer"),
            //    SalesPerson = d.Field<string>("SalesPerson")
            //}).ToList();
            //customerlst = saleslst.Select(x => new SaleP() { Customer = x.Customer, SalesPerson = x.SalesPerson }).ToList().Distinct();

        }

        private void btnBBack_Click(object sender, EventArgs e)
        {
            Form1 frm = new Form1();
            frm.Show();
            this.Hide();
        }
    }
}
