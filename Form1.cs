using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ayoti
{
    public partial class Form1 : Form
    {
        System.Data.DataTable dtProducts, dtSales;
        DataSet dsTemp;
        string EXCEL_PATH = "";
        string DBCon = ConfigurationManager.ConnectionStrings["ayoti.Properties.Settings.AyotiConnectionString"].ConnectionString;
        public Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
        public Workbook excelWorkBook;
        public Worksheet workSheet;
        string columnname = "", customer = "";
        string salesperson = "";
        int shct = 1;
        List<string> salesp = new List<string>();
        private SqlConnection con;
        private SqlCommand cmd;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txtstatus.Text = "Started";
            button1.Enabled = false;
            lblfile.Text = EXCEL_PATH;
            ReadExcelFile();
            button1.Enabled = true;
        }

        private void btnview_Click(object sender, EventArgs e)
        {
            Form frm = new ViewTableForm();
            frm.Show();
            this.Hide();
        }

        private DataColumn CreateIdentityColumn(string columnName = "")
        {
            DataColumn dc = new DataColumn(columnName);
            dc.AutoIncrement = true;
            dc.AutoIncrementSeed = dc.AutoIncrementStep = 1;
            return dc;
        }


        private void addProduct(string product, string category, int colno)
        {
            DataRow dr = dtProducts.NewRow();
            dr["Product"] = product;
            dr["Category"] = category;
            dr["ColNo"] = colno;
            dtProducts.Rows.Add(dr);
        }

        private void addSale(double quantity, string category, int colno)
        {
            if (quantity > 0)
            {
                string product = (from DataRow r in dtProducts.Rows
                                  where (int)r["ColNo"] == colno
                                  select (string)r["Product"]).FirstOrDefault();
                if (product != "")
                {
                    DataRow dr = dtSales.NewRow();
                    dr["Customer"] = customer;
                    dr["Product"] = product;
                    dr["Quantity"] = quantity;
                    dr["Category"] = category;
                    dr["SalesPerson"] = salesperson;
                    dr["Uploaded"] = 0;
                    dtSales.Rows.Add(dr);
                }
            }
        }

        private void btnclear_Click(object sender, EventArgs e)
        {
            using (con = new SqlConnection(DBCon))
            {
                con.Open();
                string query = "DROP Table Sales";
                using (SqlCommand command = new SqlCommand(query, con))
                {
                    command.ExecuteNonQuery();
                }
                query = @"CREATE TABLE [dbo].[Sales] (
                        [Id]   INT IDENTITY (1, 1) NOT NULL,
                        [Product]     VARCHAR(200) NULL,
                        [Category]    VARCHAR(100) NULL,
                        [Quantity]    FLOAT(53)    NULL,
                        [Customer]    VARCHAR(100) NULL,
                        [SalesPerson] NCHAR(100)   NULL,
                        [Uploaded] TINYINT DEFAULT((0)) NULL,
                        PRIMARY KEY CLUSTERED([Id] ASC)
                    );";
                using (SqlCommand command = new SqlCommand(query, con))
                {
                    command.ExecuteNonQuery();
                }
                txtstatus.Text = "Sales Table Truncated";
            }
        }

        private void btnGetExcel_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                EXCEL_PATH = openFileDialog1.FileName;
                lblfile.Text = EXCEL_PATH;
                button1.Enabled = true;             
            }
        }

        private void chkoverride_CheckedChanged(object sender, EventArgs e)
        {
            if (chkoverride.Checked)
                txtworksheet.ReadOnly = false;
            else
                txtworksheet.ReadOnly = true;
        }

        private void ReadExcelFile()
        {
            try
            {
                dsTemp = new DataSet();
                if (dsTemp.Tables.Count == 0)
                {
                    dtProducts = dsTemp.Tables.Add("Products");
                    dtProducts.Columns.Add(CreateIdentityColumn("Id"));
                    dtProducts.Columns.Add("Product");
                    dtProducts.Columns.Add("Category");
                    dtProducts.Columns.Add("ColNo", typeof(Int32));

                    dtSales = dsTemp.Tables.Add("Sales");
                    dtSales.Columns.Add(CreateIdentityColumn("Id"));
                    dtSales.Columns.Add("Product");
                    dtSales.Columns.Add("Category");
                    dtSales.Columns.Add("Customer");
                    dtSales.Columns.Add("SalesPerson");
                    dtSales.Columns.Add("Quantity", typeof(Double));
                    dtSales.Columns.Add("Uploaded", typeof(int));
                }

                
                //Opening/Loading the workBook in memory
                excelWorkBook = excelApp.Workbooks.Open(EXCEL_PATH);

                //retrieving the worksheet counts inside the excel workbook
                int workSheetCounts = excelWorkBook.Worksheets.Count;
                richTextBox1.Text += "workSheetCounts:" + workSheetCounts;
                //workSheetCounts = 3;
                int totalColumns = 0;
                int totalRows = 0;
                int usableRows = 0;
                int usableColumns = 0;
                Range objRange = null;
                int spiritsColStart = 0;
                int spiritsColEnd = 0;
                int beerColStart = 0;
                int beerColEnd = 0;
                int premixColStart = 0;
                int premixColEnd = 0;
                int nonalColStart = 0;
                int nonalColEnd = 0;
                int ciderColStart = 0;
                int ciderColEnd = 0;
                int mergedColumns = 0;
                int col = 0;
                if (chkoverride.Checked && txtworksheet.TextLength > 0)
                {
                    shct = int.Parse(txtworksheet.Text);
                    workSheetCounts = shct;
                }
                txtworksheetcount.Text = shct.ToString();
                for (int sheetCounter = shct; sheetCounter <= workSheetCounts; sheetCounter++)
                {
                    workSheet = excelWorkBook.Sheets[sheetCounter];
                    salesperson = workSheet.Name;
                    totalColumns = workSheet.UsedRange.Cells.Columns.Count + 1;
                    usableColumns = totalColumns - 16;
                    totalRows = workSheet.UsedRange.Cells.Rows.Count;
                    usableRows = totalRows - 37;
                    object[] data = null;
                    richTextBox1.Text += "Excel Worksheet: " + sheetCounter + ",Salesperson:" + salesperson;
                    //Iterating from row 3 because first row contains HeaderNames
                    for (int row = 3; row < usableRows; row++)
                    {
                        data = new object[usableColumns - 1];
                        if (row == 3)
                        {
                            for (col = 1; col < usableColumns; col++)
                            {
                                objRange = workSheet.Cells[row, col];
                                if (objRange.MergeCells)
                                {
                                    data[col - 1] = Convert.ToString(((Range)objRange.MergeArea[1, 1]).Text).Trim();
                                    mergedColumns = objRange.MergeArea.Columns.Count;
                                    //Debug.WriteLine("mergedColumns" + mergedColumns);
                                    //Debug.WriteLine(data[col - 1] + " row" + row + ",col" + col);
                                    columnname = data[col - 1].ToString().ToLower();
                                    switch (columnname)
                                    {
                                        case "spirits":
                                            //Debug.WriteLine(data[col - 1] + " row" + row + ",col" + col);
                                            spiritsColStart = col - 1;
                                            spiritsColEnd = col + mergedColumns - 1;
                                            //Debug.WriteLine(data[spiritsColEnd] + " spiritsColEnd" + spiritsColEnd + ",col" + col);
                                            break;
                                        case "beer":
                                            beerColStart = col - 1;
                                            beerColEnd = col + mergedColumns - 1;
                                            break;
                                        case "alcoholic pre-mix drink":
                                            premixColStart = col - 1;
                                            premixColEnd = col + mergedColumns - 1;
                                            break;
                                        case "non-alcoholic beverage":
                                            nonalColStart = col - 1;
                                            nonalColEnd = col + mergedColumns - 1;
                                            break;
                                        case "cider/perry":
                                            ciderColStart = col - 1;
                                            ciderColEnd = col + mergedColumns - 1;
                                            break;
                                    }
                                    //skip to next unmerged cell to read value
                                    col += mergedColumns - 1;

                                    //Debug.WriteLine(data[col - 1] + " row" + row + ",col" + col);
                                }
                                else
                                {
                                    data[col - 1] = Convert.ToString(objRange.Text).Trim();
                                    //Debug.WriteLine(data[col - 1] + " row si mege" + row + ",col" + col);
                                    if (data[col - 1] != null)
                                    { // cider not always a merged cell
                                        if (data[col - 1].ToString().ToLower() == "cider/perry")
                                        {
                                            ciderColStart = col - 1;
                                            ciderColEnd = col - 1;
                                        }
                                        if (data[col - 1].ToString().ToLower() == "non-alcoholic beverage")
                                        {
                                            nonalColStart = col - 1;
                                            nonalColEnd = col - 1;
                                        }

                                    }
                                }
                                //Debug.WriteLine(data[col - 1] + " row" + row + ",col" + col);
                            }
                        }
                        else
                        {
                            for (col = 1; col < usableColumns; col++)
                            {
                                objRange = workSheet.Cells[row, col];
                                if (objRange.MergeCells)
                                {
                                    data[col - 1] = Convert.ToString(((Range)objRange.MergeArea[1, 1]).Text).Trim();
                                    mergedColumns = objRange.MergeArea.Columns.Count;
                                    //Debug.WriteLine(data[col - 1] + " row" + row + ",col" + col);
                                    if (data[col - 1] != null)
                                    {
                                        string product = data[col - 1].ToString().Trim();
                                        if (product != "" && product != "Achieved") //avoid adding totals
                                        {
                                            switch (row)
                                            {
                                                case 4:
                                                    if ((col - 1) >= spiritsColStart && (col - 1) <= spiritsColEnd)
                                                    {
                                                        addProduct(product, "SPIRITS", col - 1);
                                                    }
                                                    else if ((col - 1) >= beerColStart && (col - 1) <= beerColEnd)
                                                    {
                                                        addProduct(product, "BEER", col - 1);
                                                    }
                                                    else if ((col - 1) >= premixColStart && (col - 1) <= premixColEnd)
                                                    {
                                                        addProduct(product, "PREMIX", col - 1);
                                                    }
                                                    else if ((col - 1) >= ciderColStart && (col - 1) <= ciderColEnd)
                                                    {
                                                        addProduct(product, "CIDER", col - 1);
                                                    }
                                                    else if ((col - 1) >= nonalColStart && (col - 1) <= nonalColEnd)
                                                    {
                                                        addProduct(product, "PREMIX", col - 1);
                                                    }
                                                    break;
                                                case 5:
                                                    break;
                                                default:
                                                    if ((col - 1) == 2) customer = product;

                                                    if ((col - 1) >= spiritsColStart && (col - 1) <= spiritsColEnd)
                                                    {
                                                        addSale(double.Parse(product), "SPIRIT", col - 1);
                                                    }
                                                    else if ((col - 1) >= beerColStart && (col - 1) <= beerColEnd)
                                                    {
                                                        addSale(double.Parse(product), "BEER", col - 1);
                                                    }
                                                    else if ((col - 1) >= premixColStart && (col - 1) <= premixColEnd)
                                                    {
                                                        addSale(double.Parse(product), "PREMIX", col - 1);
                                                    }
                                                    else if ((col - 1) >= ciderColStart && (col - 1) <= ciderColEnd)
                                                    {
                                                        addSale(double.Parse(product), "CIDER", col - 1);
                                                    }
                                                    else if ((col - 1) >= nonalColStart && (col - 1) <= nonalColEnd)
                                                    {
                                                        addSale(double.Parse(product), "PREMIX", col - 1);
                                                    }
                                                    break;
                                            }
                                        }
                                    }
                                    //skip to next unmerged cell to read value
                                    col += mergedColumns - 1;
                                    //Debug.WriteLine(data[col - 1] + " row" + row + ",col" + col);
                                }
                                else
                                {
                                    data[col - 1] = Convert.ToString(objRange.Text).Trim();

                                    if (data[col - 1] != null)
                                    {
                                        string product = data[col - 1].ToString().Trim();
                                        if (product != "" && product != "Achieved") //avoid adding totals
                                        {
                                            switch (row)
                                            {
                                                case 4:
                                                    if ((col - 1) >= spiritsColStart && (col - 1) <= spiritsColEnd)
                                                    {
                                                        addProduct(product, "SPIRITS", col - 1);
                                                    }
                                                    else if ((col - 1) >= beerColStart && (col - 1) <= beerColEnd)
                                                    {
                                                        addProduct(product, "BEER", col - 1);
                                                    }
                                                    else if ((col - 1) >= premixColStart && (col - 1) <= premixColEnd)
                                                    {
                                                        addProduct(product, "PREMIX", col - 1);
                                                    }
                                                    else if ((col - 1) >= ciderColStart && (col - 1) <= ciderColEnd)
                                                    {
                                                        addProduct(product, "CIDER", col - 1);
                                                    }
                                                    else if ((col - 1) >= nonalColStart && (col - 1) <= nonalColEnd)
                                                    {
                                                        addProduct(product, "PREMIX", col - 1);
                                                    }
                                                    break;
                                                case 5:
                                                    break;
                                                default:
                                                    if ((col - 1) == 2) customer = product;

                                                    if ((col - 1) >= spiritsColStart && (col - 1) <= spiritsColEnd)
                                                    {
                                                        addSale(double.Parse(product), "SPIRIT", col - 1);
                                                    }
                                                    else if ((col - 1) >= beerColStart && (col - 1) <= beerColEnd)
                                                    {
                                                        if (product.Length < 1)
                                                            Debug.WriteLine(product);
                                                        addSale(double.Parse(product), "BEER", col - 1);
                                                    }
                                                    else if ((col - 1) >= premixColStart && (col - 1) <= premixColEnd)
                                                    {
                                                        addSale(double.Parse(product), "PREMIX", col - 1);
                                                    }
                                                    else if ((col - 1) >= ciderColStart && (col - 1) <= ciderColEnd)
                                                    {
                                                        addSale(double.Parse(product), "CIDER", col - 1);
                                                    }
                                                    else if ((col - 1) >= nonalColStart && (col - 1) <= nonalColEnd)
                                                    {
                                                        addSale(double.Parse(product), "PREMIX", col - 1);
                                                    }
                                                    break;
                                            }

                                        }

                                    }
                                }
                                //if ((col - 1) >= spiritsColStart && (col - 1) <= spiritsColEnd)
                                //    Debug.WriteLine(data[col - 1] + " row" + row + ",col" + col);
                            }
                            if (row == 4) row += 1;//skip to customer
                            //Debug.WriteLine( " row" + row + ",col" + col);
                        }


                    }

                    //Debug.WriteLine("spiritsColStart" + spiritsColStart + ",spiritsColEnd" + spiritsColEnd);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                //Release the Excel objects  
                //excelWorkBook.Close(false, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
                //excelApplication.Workbooks.Close();                
                //excelApplication.Quit();
                //excelApplication = null;
                //excelWorkBook = null;
                if(excelApp!= null)
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    //Marshal.ReleaseComObject(workSheet);
                    if(excelWorkBook != null)
                    {
                        excelWorkBook.Close(false);
                        Marshal.ReleaseComObject(excelWorkBook);
                    }                    

                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                if (excelApp != null)
                {
                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }
                }


                //Marshal.ReleaseComObject(excelApp);
                //Marshal.ReleaseComObject(workSheet);

                //GC.GetTotalMemory(false);

                //GC.GetTotalMemory(true);

                DumpDataIntoSql();
            }
        }
        
        private void DumpDataIntoSql()
        {
            try
            {
                dtProducts = null;
                int rowsAffected = 0;
                richTextBox1.Text += "Dumping Data into SQL Tables\n";
                con = new SqlConnection(DBCon);
                con.Open();
                string query = string.Empty;

                rowsAffected = 0;
                string colNames = string.Join(",", dsTemp.Tables[1].Columns.Cast<DataColumn>().Where(x => x.AutoIncrement == false).Select(x => x.ColumnName).ToArray<string>());
                string[] arr = colNames.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                query = "insert into " + dsTemp.Tables[1].TableName + "(" + colNames + ") values(";
                for (int col = 0; col < arr.Length; col++)
                {
                    if (col != arr.Length - 1)
                        query += "@" + arr[col] + ",";
                    else
                        query += "@" + arr[col] + ")";
                }
                cmd = new SqlCommand(query, con);
                for (int row = 0; row < dsTemp.Tables[1].Rows.Count; row++)
                {
                    for (int col = 0, arrCounter = 0; col < dsTemp.Tables[1].Columns.Count; col++)
                    {
                        if (!dsTemp.Tables[1].Columns[col].AutoIncrement)
                        {
                            cmd.Parameters.AddWithValue("@" + arr[arrCounter], dsTemp.Tables[1].Rows[row][col].ToString());
                            arrCounter++;
                        }
                    }
                    rowsAffected += cmd.ExecuteNonQuery();
                    cmd.Parameters.Clear();
                }
                dtSales = null;
                dsTemp = null;
                query = "DELETE FROM Sales WHERE Product = ' '";
                int rowsDeleted = 0;
                using (SqlCommand command = new SqlCommand(query, con))
                {
                    rowsDeleted = command.ExecuteNonQuery();
                }
                con.Close();
                rowsAffected = rowsAffected - rowsDeleted;
                richTextBox1.Text += rowsAffected  + "Records Affected For sales";

                
                

                txtrows.Text = rowsAffected.ToString();
                txtstatus.Text = "Completed";


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
