using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using Syncfusion.XlsIO;
namespace DCH_OpenBalance
{
    public partial class DCH_OpenBalance : Form
    {
      
        public DCH_OpenBalance()
        {
            InitializeComponent();
        }
        #region Variable
        DataTable dtexcel = new DataTable();
        DataTable dt = new DataTable();
        int ProgressMaximum = 0;
        //Microsoft.Office.Interop.Excel.Application excel;
        //Microsoft.Office.Interop.Excel.Workbook excelworkBook;
        //Microsoft.Office.Interop.Excel.Worksheet excelSheet;
        //Microsoft.Office.Interop.Excel.Range excelCellrange;
        #endregion

        public class ETMItemBlance
        {
            public string Warehouse_ID { get; set; }
            public string Location_Code { get; set; }
            public string LPN { get; set; }
            public string Item_Client_ID { get; set; }
            public string Item_Number { get; set; }
            public string Lot_Number { get; set; }
            public string Supplier_Lot_Number { get; set; }
            public string Received_Date { get; set; }
            public string Manufactured_Date { get; set; }
            public string Expiration_Date { get; set; }
            public string Base_Unit_Qty { get; set; }
            public string Base_Unit_UOM { get; set; }
            public string Inventory_Status { get; set; }
            public string Attribute_1 { get; set; }
            public string Attribute_2 { get; set; }
            public string Attribute_3 { get; set; }
            public string Attribute_4 { get; set; }
            public string sItem_Number { get; set; }
            public string sLot_Number { get; set; }
            public string sExpire_date { get; set; }
            public string sAttribute1 { get; set; }
            public string sAttribute2 { get; set; }
            public string sAttribute3 { get; set; }
            public string sAttribute4 { get; set; }
            public string sUOM { get; set; }
            public string sBase_Qty { get; set; }
            public string sInventory_Status { get; set; }
            public string sCheck_Expire_date { get; set; }
            public string sCheckReceived_Date { get; set; }
            public string sCheckManufactured_Date { get; set; }
        }
        private DataTable AddColumnDatatable()
        {
            DataTable dt = new DataTable();
            try
            {
                dt.Columns.Add("Warehouse_ID");
                dt.Columns.Add("Location_Code");
                dt.Columns.Add("LPN");
                dt.Columns.Add("Item_Client_ID");
                dt.Columns.Add("Item_Number");
                dt.Columns.Add("Lot_Number");
                dt.Columns.Add("Supplier_Lot_Number");
                dt.Columns.Add("Received_Date");
                dt.Columns.Add("Manufactured_Date");
                dt.Columns.Add("Expiration_Date");
                dt.Columns.Add("Base_Unit_Qty");
                dt.Columns.Add("Base_Unit_UOM");
                dt.Columns.Add("Inventory_Status");
                dt.Columns.Add("Attribute_1");
                dt.Columns.Add("Attribute_2");
                dt.Columns.Add("Attribute_3");
                dt.Columns.Add("Attribute_4");
                return dt;
            }
            catch (Exception ex)
            {
                return dt;
            }
        }
        private void DCH_OpenBalance_Load(object sender, EventArgs e)
        {
            dt = new DataTable();
            dt = AddColumnDatatable();
            dataGridView_Display.Visible = true;
            dataGridView_Display.DataSource = dt;
        }
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            dtexcel = new DataTable();
            dt = new DataTable();
            // Set cursor as hourglass
            Cursor.Current = Cursors.WaitCursor;
            //int i;
          
            try
            {
                string conn = string.Empty;
                if (fileExt.CompareTo(".xls") == 0)
                    conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
                else
                    conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
                using (OleDbConnection con = new OleDbConnection(conn))
                {
                    try
                    {

                        OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [OPENING$]", con); //here we read data from sheet1  
                        oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                        dt = new DataTable();
                        dt = AddColumnDatatable();
                        for (int i = 2; i < dtexcel.Rows.Count; i++)
                        {
                            DataRow dr = dt.NewRow();
                            dr["Warehouse_ID"] = dtexcel.Rows[i]["F1"].ToString();
                            dr["Location_Code"] = dtexcel.Rows[i]["F2"].ToString();
                            dr["LPN"] = dtexcel.Rows[i]["F3"].ToString();
                            dr["Item_Client_ID"] = dtexcel.Rows[i]["F4"].ToString();
                            dr["Item_Number"] = dtexcel.Rows[i]["F5"].ToString();
                            dr["Lot_Number"] = dtexcel.Rows[i]["F6"].ToString();
                            dr["Supplier_Lot_Number"] = dtexcel.Rows[i]["F7"].ToString();
                            dr["Received_Date"] = dtexcel.Rows[i]["F8"].ToString();
                            dr["Manufactured_Date"] = dtexcel.Rows[i]["F9"].ToString();
                            dr["Expiration_Date"] = dtexcel.Rows[i]["F10"].ToString();
                            dr["Base_Unit_Qty"] = dtexcel.Rows[i]["F11"].ToString();
                            dr["Base_Unit_UOM"] = dtexcel.Rows[i]["F12"].ToString();
                            dr["Inventory_Status"] = dtexcel.Rows[i]["F13"].ToString();
                            dr["Attribute_1"] = dtexcel.Rows[i]["F14"].ToString();
                            dr["Attribute_2"] = dtexcel.Rows[i]["F15"].ToString();
                            dr["Attribute_3"] = dtexcel.Rows[i]["F16"].ToString();
                            dr["Attribute_4"] = dtexcel.Rows[i]["F17"].ToString();
                            dt.Rows.Add(dr);
                        }
                    }
                    catch (Exception ex)
                    {
                        return dt;
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {
                return dt;
            }
            finally
            {
                // Set cursor as default arrow
                Cursor.Current = Cursors.Default;
            }
           
        }
        private void Button_Choose_Click(object sender, EventArgs e)
        {
            try
            {
                string filePath = string.Empty;
                string fileExt = string.Empty;
                OpenFileDialog file = new OpenFileDialog(); //open dialog to choose file  
                if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK) //if there is a file choosen by the user  
                {
                    filePath = file.FileName; //get the path of the file  
                    fileExt = Path.GetExtension(filePath); //get the file extension  
                    if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                    {
                        try
                        {
                            DataTable dtExcel = new DataTable();
                            dtExcel = ReadExcel(filePath, fileExt); //read excel file  
                            dataGridView_Display.Visible = true;
                            dataGridView_Display.DataSource = dtExcel;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error); //custom messageBox to show error  
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
           
        }
        private bool ExecuteSqlTransaction(DataTable dt)
        {
            string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
            bool bResult = false;
        
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                SqlTransaction transaction;
                transaction = connection.BeginTransaction();
                //using (SqlTransaction sqlTrans = connection.BeginTransaction())

                DataTable dt_count = new DataTable();
                string sql_count = @"SELECT case when max(nRound) + 1 is null then '1' else max(nRound) + 1 end nRound FROM TMItemBlance with(nolock)";
                try
                {
                    SqlDataAdapter da = new SqlDataAdapter(sql_count, connectionString);
                    da.Fill(dt_count);
                    //using (sqlTrans = connection.BeginTransaction())
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sql = @"INSERT INTO  TMItemBlance
                                       (Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number
                                       ,Supplier_Lot_Number,Received_Date,Manufactured_Date,Expiration_Date
                                       ,Base_Unit_Qty,Base_Unit_UOM,Inventory_Status,Attribute_1,Attribute_2
                                       ,Attribute_3,Attribute_4,nRound)
                                        VALUES
                                       (@Warehouse_ID,@Location_Code,@LPN,@Item_Client_ID,@Item_Number,@Lot_Number
                                       ,@Supplier_Lot_Number,@Received_Date,@Manufactured_Date,@Expiration_Date
                                       ,@Base_Unit_Qty,@Base_Unit_UOM,@Inventory_Status,@Attribute_1,@Attribute_2
                                       ,@Attribute_3,@Attribute_4,@nRound)";

                   
                        using (SqlCommand sqlCommand = new SqlCommand(sql, connection, transaction))
                        {
                            sqlCommand.Parameters.AddWithValue("@Warehouse_ID", dt.Rows[i]["Warehouse_ID"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Location_Code", dt.Rows[i]["Location_Code"].ToString());
                            sqlCommand.Parameters.AddWithValue("@LPN", dt.Rows[i]["LPN"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Item_Client_ID", dt.Rows[i]["Item_Client_ID"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Item_Number", dt.Rows[i]["Item_Number"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Lot_Number", dt.Rows[i]["Lot_Number"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Supplier_Lot_Number", dt.Rows[i]["Supplier_Lot_Number"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Received_Date", dt.Rows[i]["Received_Date"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Manufactured_Date", dt.Rows[i]["Manufactured_Date"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Expiration_Date", dt.Rows[i]["Expiration_Date"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Base_Unit_Qty", dt.Rows[i]["Base_Unit_Qty"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Base_Unit_UOM", dt.Rows[i]["Base_Unit_UOM"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Inventory_Status", dt.Rows[i]["Inventory_Status"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Attribute_1", dt.Rows[i]["Attribute_1"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Attribute_2", dt.Rows[i]["Attribute_2"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Attribute_3", dt.Rows[i]["Attribute_3"].ToString());
                            sqlCommand.Parameters.AddWithValue("@Attribute_4", dt.Rows[i]["Attribute_4"].ToString());
                            sqlCommand.Parameters.AddWithValue("@nRound", dt_count.Rows[0]["nRound"].ToString());
                            sqlCommand.ExecuteNonQuery();
                        }
                    }

                    transaction.Commit();
                    bResult = true;
                    return bResult;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    transaction.Rollback();
                    bResult = false;
                    return bResult;

                }
                finally
                {
                    connection.Close();
                }
            }
        }
        private bool Save_TMItemBlance_Pass(ETMItemBlance lst)
        {
            string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
            bool bResult = false;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                SqlTransaction transaction;
                transaction = connection.BeginTransaction();
                try
                {

                    string sql_delete = @"Select * from TMItemBlance_Pass with(nolock)  ";
                    sql_delete += " Where Warehouse_ID = '"+ lst.Warehouse_ID.ToString().Trim() + "' and Location_Code = '" + lst.Location_Code.ToString().Trim() + "'";
                    sql_delete += " and LPN = '" + lst.LPN.ToString().Trim() + "' and Item_Client_ID = '" + lst.Item_Client_ID.ToString().Trim() + "'  ";
                    sql_delete += " and Item_Number = '" + lst.Item_Number.ToString().Trim() + "' and Lot_Number = '" + lst.Lot_Number.ToString().Trim() + "' ";
                    DataTable dt = new DataTable();
                    SqlDataAdapter da = new SqlDataAdapter(sql_delete, connectionString);
                    da.Fill(dt);
                    if (dt.Rows.Count != 0)
                    {
                        string sql_delete1 = @"Delete from TMItemBlance_Pass
                                       Where Warehouse_ID = @Warehouse_ID and Location_Code = @Location_Code and LPN = @LPN and Item_Client_ID = @Item_Client_ID and 
                                       Item_Number = @Item_Number and Lot_Number = @Lot_Number";


                        using (SqlCommand sqlCommand = new SqlCommand(sql_delete1, connection, transaction))
                        {
                            sqlCommand.Parameters.AddWithValue("@Warehouse_ID", lst.Warehouse_ID.ToString().Trim());
                            sqlCommand.Parameters.AddWithValue("@Location_Code", lst.Location_Code.ToString().Trim());
                            sqlCommand.Parameters.AddWithValue("@LPN", lst.LPN.ToString().Trim());
                            sqlCommand.Parameters.AddWithValue("@Item_Client_ID", lst.Item_Client_ID.ToString().Trim());
                            sqlCommand.Parameters.AddWithValue("@Item_Number", lst.Item_Number.ToString().Trim());
                            sqlCommand.Parameters.AddWithValue("@Lot_Number", lst.Lot_Number.ToString().Trim());
                            sqlCommand.ExecuteNonQuery();
                        }

                    }

                    string sql = @"INSERT INTO  TMItemBlance_Pass
                                       (Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number
                                       ,Supplier_Lot_Number,Received_Date,Manufactured_Date,Expiration_Date
                                       ,Base_Unit_Qty,Base_Unit_UOM,Inventory_Status,Attribute_1,Attribute_2
                                       ,Attribute_3,Attribute_4)
                                        VALUES
                                       (@Warehouse_ID,@Location_Code,@LPN,@Item_Client_ID,@Item_Number,@Lot_Number
                                       ,@Supplier_Lot_Number,@Received_Date,@Manufactured_Date,@Expiration_Date
                                       ,@Base_Unit_Qty,@Base_Unit_UOM,@Inventory_Status,@Attribute_1,@Attribute_2
                                       ,@Attribute_3,@Attribute_4)";


                    using (SqlCommand sqlCommand = new SqlCommand(sql, connection, transaction))
                    {
                        sqlCommand.Parameters.AddWithValue("@Warehouse_ID", lst.Warehouse_ID.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Location_Code", lst.Location_Code.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@LPN", lst.LPN.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Item_Client_ID", lst.Item_Client_ID.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Item_Number", lst.Item_Number.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Lot_Number", lst.Lot_Number.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Supplier_Lot_Number", lst.Supplier_Lot_Number.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Received_Date", lst.Received_Date.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Manufactured_Date", lst.Manufactured_Date.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Expiration_Date", lst.Expiration_Date.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Base_Unit_Qty", lst.Base_Unit_Qty.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Base_Unit_UOM", lst.Base_Unit_UOM.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Inventory_Status", lst.Inventory_Status.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Attribute_1", lst.Attribute_1.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Attribute_2", lst.Attribute_2.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Attribute_3", lst.Attribute_3.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Attribute_4", lst.Attribute_4.ToString().Trim());
                        sqlCommand.ExecuteNonQuery();
                    }

                    transaction.Commit();
                    bResult = true;
                    return bResult;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    transaction.Rollback();
                    bResult = false;
                    return bResult;

                }
                finally
                {
                    connection.Close();
                }
            }
        }
        private void Show_Valid()
        {
            Cursor.Current = Cursors.WaitCursor;
            string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
            try
            {
                DataTable dt_valid = new DataTable();
                string sql_display = @" SELECT Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number,Supplier_Lot_Number
,Received_Date,Manufactured_Date,Expiration_Date,Base_Unit_Qty,Base_Unit_UOM,Inventory_Status,Attribute_1
,Attribute_2,Attribute_3,Attribute_4,sItem_Number,sLot_Number,sExpire_date,sAttribute1,sAttribute2
,sAttribute3,sAttribute4,sInventory_Status,sUOM,sBase_Qty,sCheck_Expire_date,CheckReceived_Date,CheckManufactured_Date
FROM  TMItemBlance_Valid with(nolock)";

                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_valid);
                dataGridView_Display.Visible = true;
                dataGridView_Display.DataSource = dt_valid;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }
        private void Show_Pass()
        {
            Cursor.Current = Cursors.WaitCursor;
            string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
            try
            {
                DataTable dt_pass = new DataTable();
                string sql_display = @" SELECT Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number,Supplier_Lot_Number
,Received_Date,Manufactured_Date,Expiration_Date,Base_Unit_Qty,Base_Unit_UOM,Inventory_Status,Attribute_1
,Attribute_2,Attribute_3,Attribute_4
FROM  TMItemBlance_Pass with(nolock)";

                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_pass);
                dataGridView_Display.Visible = true;
                dataGridView_Display.DataSource = dt_pass;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }
        private bool Delete_TMItemBlance_Valid()
        {
            string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                SqlTransaction transaction;
                transaction = connection.BeginTransaction();
                try
                {
                    string sql = "Delete from TMItemBlance_Valid";
                    using (SqlCommand sqlCommand = new SqlCommand(sql, connection, transaction))
                    {
                        sqlCommand.ExecuteNonQuery();
                    }
                    transaction.Commit();
                    return true;
                }
                catch (Exception)
                {
                    transaction.Rollback();
                    return false;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        private bool Save_TMItemBlance_Valid(ETMItemBlance lst)
        {
            string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
            bool bResult = false;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                SqlTransaction transaction;
                transaction = connection.BeginTransaction();
                try
                {

                    string sql = @"INSERT INTO  TMItemBlance_Valid
                                       (Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number
                                       ,Supplier_Lot_Number,Received_Date,Manufactured_Date,Expiration_Date
                                       ,Base_Unit_Qty,Base_Unit_UOM,Inventory_Status,Attribute_1,Attribute_2
                                       ,Attribute_3,Attribute_4,sItem_Number,sLot_Number,sExpire_date,sAttribute1,sAttribute2,sAttribute3,sAttribute4,sInventory_Status,sUOM,sBase_Qty,sCheck_Expire_date,CheckReceived_Date,CheckManufactured_Date)
                                        VALUES
                                       (@Warehouse_ID,@Location_Code,@LPN,@Item_Client_ID,@Item_Number,@Lot_Number
                                       ,@Supplier_Lot_Number,@Received_Date,@Manufactured_Date,@Expiration_Date
                                       ,@Base_Unit_Qty,@Base_Unit_UOM,@Inventory_Status,@Attribute_1,@Attribute_2
                                       ,@Attribute_3,@Attribute_4,@sItem_Number,@sLot_Number,@sExpire_date,@sAttribute1,@sAttribute2,@sAttribute3,@sAttribute4,@sInventory_Status,@sUOM,@sBase_Qty,@sCheck_Expire_date,@CheckReceived_Date,@CheckManufactured_Date)";


                    using (SqlCommand sqlCommand = new SqlCommand(sql, connection, transaction))
                    {
                        sqlCommand.Parameters.AddWithValue("@Warehouse_ID", lst.Warehouse_ID.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Location_Code", lst.Location_Code.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@LPN", lst.LPN.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Item_Client_ID", lst.Item_Client_ID.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Item_Number", lst.Item_Number.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Lot_Number", lst.Lot_Number.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Supplier_Lot_Number", lst.Supplier_Lot_Number.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Received_Date", lst.Received_Date.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Manufactured_Date", lst.Manufactured_Date.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Expiration_Date", lst.Expiration_Date.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Base_Unit_Qty", lst.Base_Unit_Qty.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Base_Unit_UOM", lst.Base_Unit_UOM.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Inventory_Status", lst.Inventory_Status.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Attribute_1", lst.Attribute_1.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Attribute_2", lst.Attribute_2.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Attribute_3", lst.Attribute_3.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@Attribute_4", lst.Attribute_4.ToString().Trim());

                        sqlCommand.Parameters.AddWithValue("@sItem_Number", lst.sItem_Number.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sLot_Number", lst.sLot_Number.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sExpire_date", lst.sExpire_date.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sAttribute1", lst.sAttribute1.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sAttribute2", lst.sAttribute2.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sAttribute3", lst.sAttribute3.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sAttribute4", lst.sAttribute4.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sInventory_Status", lst.sInventory_Status.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sUOM", lst.sUOM.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sBase_Qty", lst.sBase_Qty.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@sCheck_Expire_date", lst.sCheck_Expire_date.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@CheckReceived_Date", lst.sCheckReceived_Date.ToString().Trim());
                        sqlCommand.Parameters.AddWithValue("@CheckManufactured_Date", lst.sCheckManufactured_Date.ToString().Trim());
                        sqlCommand.ExecuteNonQuery();
                    }
                    transaction.Commit();
                    bResult = true;
                    return bResult;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    transaction.Rollback();
                    bResult = false;
                    return bResult;

                }
                finally
                {
                    connection.Close();
                }
            }
        }
        private void Button_Save_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count != 0)
            {
                Cursor.Current = Cursors.WaitCursor;
                try
                {
                    if (MessageBox.Show("Confirm Save Data ?", "Dialog Box Confirm", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                    {

                        bool bResut = ExecuteSqlTransaction(dt);
                        if (bResut)
                        {
                            dt = new DataTable();
                            dt = AddColumnDatatable();
                            dataGridView_Display.Visible = true;
                            dataGridView_Display.DataSource = dt;
                            MessageBox.Show("Success");
                        }
                        else // 
                        {
                            dt = new DataTable();
                            dt = AddColumnDatatable();
                            dataGridView_Display.Visible = true;
                            dataGridView_Display.DataSource = dt;
                            MessageBox.Show("Not success");
                        }
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    return;
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }
            }
            else
            {
                MessageBox.Show("Please click Choose and Read File button before");
            }
           
        }
        private void Button_Display_Click(object sender, EventArgs e)
        {
            Show_Valid();
        }
        private void Export_Excel(DataTable dt)
        {
            try
            {
                //Create an instance of ExcelEngine
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    //Initialize Application
                    IApplication application = excelEngine.Excel;

                    //Set the default application version as Excel 2016
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    //Create a new workbook
                    IWorkbook workbook = application.Workbooks.Create(1);


                    //Access first worksheet from the workbook instance
                    IWorksheet worksheet = workbook.Worksheets[0];
                    //worksheet.Range["A1:Q1"].Text = string.Empty;
                    worksheet.Name = "OPENING";
                    //Exporting DataTable to worksheet
                    DataTable dataTable = dt;
                    //worksheet.ImportDataTable(dataTable, true, 1, 1);

                    int ncount = 0;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        ncount = ncount + 1;
                        worksheet.Range["A"+ ncount.ToString()].Text = dt.Rows[i]["Warehouse_ID"].ToString();
                        worksheet.Range["B" + ncount.ToString()].Text = dt.Rows[i]["Location_Code"].ToString();
                        worksheet.Range["C" + ncount.ToString()].Text = dt.Rows[i]["LPN"].ToString();
                        worksheet.Range["D" + ncount.ToString()].Text = dt.Rows[i]["Item_Client_ID"].ToString();
                        worksheet.Range["E" + ncount.ToString()].Text = dt.Rows[i]["Item_Number"].ToString();
                        worksheet.Range["F" + ncount.ToString()].Text = dt.Rows[i]["Lot_Number"].ToString();
                        worksheet.Range["G" + ncount.ToString()].Text = dt.Rows[i]["Supplier_Lot_Number"].ToString();
                        worksheet.Range["H" + ncount.ToString()].Text = dt.Rows[i]["Received_Date"].ToString();
                        worksheet.Range["I" + ncount.ToString()].Text = dt.Rows[i]["Manufactured_Date"].ToString();
                        worksheet.Range["J" + ncount.ToString()].Text = dt.Rows[i]["Expiration_Date"].ToString();
                        worksheet.Range["K" + ncount.ToString()].Text = dt.Rows[i]["Base_Unit_Qty"].ToString();
                        worksheet.Range["L" + ncount.ToString()].Text = dt.Rows[i]["Base_Unit_UOM"].ToString();
                        worksheet.Range["M" + ncount.ToString()].Text = dt.Rows[i]["Inventory_Status"].ToString();
                        worksheet.Range["N" + ncount.ToString()].Text = dt.Rows[i]["Attribute_1"].ToString();
                        worksheet.Range["O" + ncount.ToString()].Text = dt.Rows[i]["Attribute_2"].ToString();
                        worksheet.Range["P" + ncount.ToString()].Text = dt.Rows[i]["Attribute_3"].ToString();
                        worksheet.Range["Q" + ncount.ToString()].Text = dt.Rows[i]["Attribute_4"].ToString();
                    }

                  
                    worksheet.UsedRange.AutofitColumns();
                    worksheet.Range["A1:Q1"].CellStyle.Color = Color.Aqua;
                    worksheet.Range["A2:Q2"].CellStyle.Color = Color.Bisque;

                    //Save the workbook to disk in xlsx format
                    string sPath = textBox_Path.Text.Trim() + @"\" + textBox_Name.Text.Trim() + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + @".xlsx";
                    workbook.SaveAs(sPath);
                }
                MessageBox.Show("Export success");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        private void Export_Excel_Valid(DataTable dt)
        {
            try
            {
                //Create an instance of ExcelEngine
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    //Initialize Application
                    IApplication application = excelEngine.Excel;

                    //Set the default application version as Excel 2016
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    //Create a new workbook
                    IWorkbook workbook = application.Workbooks.Create(1);


                    //Access first worksheet from the workbook instance
                    IWorksheet worksheet = workbook.Worksheets[0];
                    //worksheet.Range["A1:Q1"].Text = string.Empty;
                    worksheet.Name = "OPENING";
                    //Exporting DataTable to worksheet
                    DataTable dataTable = dt;
                    //worksheet.ImportDataTable(dataTable, true, 1, 1);

                    int ncount = 0;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        ncount = ncount + 1;
                        worksheet.Range["A" + ncount.ToString()].Text = dt.Rows[i]["Warehouse_ID"].ToString();
                        worksheet.Range["B" + ncount.ToString()].Text = dt.Rows[i]["Location_Code"].ToString();
                        worksheet.Range["C" + ncount.ToString()].Text = dt.Rows[i]["LPN"].ToString();
                        worksheet.Range["D" + ncount.ToString()].Text = dt.Rows[i]["Item_Client_ID"].ToString();
                        worksheet.Range["E" + ncount.ToString()].Text = dt.Rows[i]["Item_Number"].ToString();
                        worksheet.Range["F" + ncount.ToString()].Text = dt.Rows[i]["Lot_Number"].ToString();
                        worksheet.Range["G" + ncount.ToString()].Text = dt.Rows[i]["Supplier_Lot_Number"].ToString();
                        worksheet.Range["H" + ncount.ToString()].Text = dt.Rows[i]["Received_Date"].ToString();
                        worksheet.Range["I" + ncount.ToString()].Text = dt.Rows[i]["Manufactured_Date"].ToString();
                        worksheet.Range["J" + ncount.ToString()].Text = dt.Rows[i]["Expiration_Date"].ToString();
                        worksheet.Range["K" + ncount.ToString()].Text = dt.Rows[i]["Base_Unit_Qty"].ToString();
                        worksheet.Range["L" + ncount.ToString()].Text = dt.Rows[i]["Base_Unit_UOM"].ToString();
                        worksheet.Range["M" + ncount.ToString()].Text = dt.Rows[i]["Inventory_Status"].ToString();
                        worksheet.Range["N" + ncount.ToString()].Text = dt.Rows[i]["Attribute_1"].ToString();
                        worksheet.Range["O" + ncount.ToString()].Text = dt.Rows[i]["Attribute_2"].ToString();
                        worksheet.Range["P" + ncount.ToString()].Text = dt.Rows[i]["Attribute_3"].ToString();
                        worksheet.Range["Q" + ncount.ToString()].Text = dt.Rows[i]["Attribute_4"].ToString();

                        //worksheet.Range["R" + ncount.ToString()].Text = dt.Rows[i]["CheckItem_Number"].ToString();
                        //worksheet.Range["S" + ncount.ToString()].Text = dt.Rows[i]["CheckLot_Number"].ToString();                        
                        //worksheet.Range["T" + ncount.ToString()].Text = dt.Rows[i]["CheckExpire_date"].ToString(); 
                        //worksheet.Range["U" + ncount.ToString()].Text = dt.Rows[i]["CheckAttribute1"].ToString();
                        //worksheet.Range["V" + ncount.ToString()].Text = dt.Rows[i]["CheckAttribute2"].ToString();
                        //worksheet.Range["W" + ncount.ToString()].Text = dt.Rows[i]["CheckAttribute3"].ToString();
                        //worksheet.Range["X" + ncount.ToString()].Text = dt.Rows[i]["CheckAttribute4"].ToString();
                        //worksheet.Range["Y" + ncount.ToString()].Text = dt.Rows[i]["CheckInventory_Status"].ToString();                      
                        //worksheet.Range["Z" + ncount.ToString()].Text = dt.Rows[i]["CheckUOM"].ToString();                    
                        //worksheet.Range["AA" + ncount.ToString()].Text = dt.Rows[i]["Check_Base_Qty"].ToString();
                        //worksheet.Range["AB" + ncount.ToString()].Text = dt.Rows[i]["Check_Max_Expire_date"].ToString();
                        //worksheet.Range["AC" + ncount.ToString()].Text = dt.Rows[i]["CheckReceived_Date"].ToString();
                        //worksheet.Range["AD" + ncount.ToString()].Text = dt.Rows[i]["CheckManufactured_Date"].ToString();

                        worksheet.Range["R" + ncount.ToString()].Text = dt.Rows[i]["CheckItem_Number"].ToString();
                        worksheet.Range["S" + ncount.ToString()].Text = dt.Rows[i]["CheckLot_Number"].ToString();
                        worksheet.Range["T" + ncount.ToString()].Text = dt.Rows[i]["CheckExpire_date"].ToString();
                        worksheet.Range["U" + ncount.ToString()].Text = dt.Rows[i]["CheckAttribute1"].ToString();
                        worksheet.Range["V" + ncount.ToString()].Text = dt.Rows[i]["CheckAttribute2"].ToString();
                        worksheet.Range["W" + ncount.ToString()].Text = dt.Rows[i]["CheckAttribute3"].ToString();
                        worksheet.Range["X" + ncount.ToString()].Text = dt.Rows[i]["CheckAttribute4"].ToString();
                        worksheet.Range["Y" + ncount.ToString()].Text = dt.Rows[i]["CheckInventory_Status"].ToString();
                        worksheet.Range["Z" + ncount.ToString()].Text = dt.Rows[i]["CheckUOM"].ToString();
                        worksheet.Range["AA" + ncount.ToString()].Text = dt.Rows[i]["Check_Base_Qty"].ToString();
                        worksheet.Range["AB" + ncount.ToString()].Text = dt.Rows[i]["Check_Max_Expire_date"].ToString();
                        worksheet.Range["AC" + ncount.ToString()].Text = dt.Rows[i]["CheckReceived_Date"].ToString();
                        worksheet.Range["AD" + ncount.ToString()].Text = dt.Rows[i]["CheckManufactured_Date"].ToString();


                        if (i >= 2)
                        {
                            if (dt.Rows[i]["CheckItem_Number"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["E" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckLot_Number"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["F" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["Check_Max_Expire_date"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["J" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["Check_Base_Qty"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["K" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckUOM"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["L" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckInventory_Status"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["M" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckAttribute4"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["Q" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckAttribute3"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["P" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckAttribute2"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["O" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckAttribute1"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["N" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckExpire_date"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["J" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckReceived_Date"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["H" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }

                            if (dt.Rows[i]["CheckManufactured_Date"].ToString().Trim() != string.Empty)
                            {
                                worksheet.Range["I" + ncount.ToString()].CellStyle.Color = Color.Red;
                            }
                        }
                    }

                     
                    worksheet.UsedRange.AutofitColumns();

                    //Save the workbook to disk in xlsx format
                    string sPath = textBox_Path.Text.Trim() + @"\" + textBox_Name.Text.Trim() + "_" + DateTime.Now.ToString("yyyyMMddhhmmss") + @".xlsx";

                    worksheet.Range["R1:AD1"].CellStyle.Color = Color.Yellow;
                    worksheet.Range["R2:AD2"].CellStyle.Color = Color.Yellow;

                    worksheet.Range["A1:Q1"].CellStyle.Color = Color.Aqua;
                    worksheet.Range["A2:Q2"].CellStyle.Color = Color.Bisque;
                    workbook.SaveAs(sPath);
                }
                MessageBox.Show("Export success");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        private void Button_ExportExcel_Click(object sender, EventArgs e)
        {
            if (textBox_Path.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please select path before");
                return;
            }

            if (textBox_Name.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please input file name before");
                return;
            }

            if (radioButton_success.Checked)
            {
                Cursor.Current = Cursors.WaitCursor;
                string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
                try
                {

                    DataTable dt_ex = new DataTable();
                    string sql_display = @"SELECT Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number,Supplier_Lot_Number,Received_Date,Manufactured_Date
                                           ,Expiration_Date,Base_Unit_Qty,Base_Unit_UOM,Inventory_Status,Attribute_1,Attribute_2,Attribute_3,Attribute_4 
                                            FROM  VTMItemBlance_ExPass with(nolock) order by nID"; // 
                    SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                    da.Fill(dt_ex);
                    if (dt_ex.Rows.Count > 2)
                    {
                        Export_Excel(dt_ex);
                    }
                    else
                    {
                        MessageBox.Show("Not found data");
                        return;
                    }

                    //textBox_Name.Text = string.Empty;
                    textBox_Path.Text = string.Empty;
                    dt = new DataTable();
                    dt = AddColumnDatatable();
                    dataGridView_Display.Visible = true;
                    dataGridView_Display.DataSource = dt_ex;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    return;
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }
            }
            else // Export Table TMMaster_Validate
            {
                Cursor.Current = Cursors.WaitCursor;
                string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
                try
                {

                    DataTable dt_ex = new DataTable();
                    string sql_display = @"SELECT Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number,Supplier_Lot_Number,Received_Date,Manufactured_Date
                                           ,Expiration_Date,Base_Unit_Qty,Base_Unit_UOM,Inventory_Status,Attribute_1,Attribute_2,Attribute_3,Attribute_4,CheckItem_Number
                                           ,CheckLot_Number,CheckExpire_date,CheckAttribute1,CheckAttribute2,CheckAttribute3,CheckAttribute4,CheckInventory_Status
                                           ,CheckUOM,Check_Base_Qty,Check_Max_Expire_date,CheckReceived_Date,CheckManufactured_Date
                                           FROM  VTMHeader_Valid with(nolock) order by nID";
                    SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                    da.Fill(dt_ex);
                    if (dt_ex.Rows.Count > 2)
                    {
                        Export_Excel_Valid(dt_ex);
                    }
                    else
                    {
                        MessageBox.Show("Not found data");
                        return;
                    }

                    //textBox_Name.Text = string.Empty;
                    textBox_Path.Text = string.Empty;
                    dt = new DataTable();
                    dt = AddColumnDatatable();
                    dataGridView_Display.Visible = true;
                    dataGridView_Display.DataSource = dt_ex;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                    return;
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }
            }
        }
        private void Button_Browse_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            folderDlg.ShowNewFolderButton = true;
            // Show the FolderBrowserDialog.  
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox_Path.Text = folderDlg.SelectedPath;
                //Environment.SpecialFolder root = folderDlg.RootFolder;
            }
        }

        #region Validate data
        private bool Valdate_Item_Number(string Item_Number, string connectionString, string Item_Client_ID)
        {
            DataTable dt_check = new DataTable();
            try
            {
                string sql_display = @" SELECT  Item_Number,bLot_Number,bExpire_date,bAttribute1,bAttribute2,bAttribute3,bAttribute4,UOM,Item_Client_ID
                                       FROM  TMMaster_Validate with(nolock) ";
                sql_display += " Where Item_Number = '" + Item_Number + "' and Item_Client_ID = '"+ Item_Client_ID + "'";
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_check);
                if (dt_check.Rows.Count != 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private bool Valdate_UOM(string Item_Number, string UOM, string connectionString, out string sError, string Item_Client_ID)
        {
            DataTable dt_check = new DataTable();
            sError = string.Empty;
            try
            {
                string sql_display = @" SELECT  Item_Number,bLot_Number,bExpire_date,bAttribute1,bAttribute2,bAttribute3,bAttribute4,UOM,Item_Client_ID
                                       FROM  TMMaster_Validate with(nolock) ";
                sql_display += " Where Item_Number = '" + Item_Number + "' and Item_Client_ID = '" + Item_Client_ID + "'";
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_check);
                if (dt_check.Rows.Count != 0)
                {
                    if (UOM.Trim() == dt_check.Rows[0]["UOM"].ToString().Trim())
                    {
                        return true;
                    }
                    else
                    {
                        sError = "UOM ไม่ตรงกับ Master ที่ถูกต้องคือ UOM : " + dt_check.Rows[0]["UOM"].ToString().Trim();
                        return false;
                    }
                }
                else
                {
                    sError = "ไม่เจอ Item Number ในระบบ";
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private bool Valdate_Base_Unit_Qty(string Base_Unit_Qty, out string sError)
        {
            sError = string.Empty;
            try
            {
                if (Convert.ToInt32(Base_Unit_Qty.Trim()) > 0)
                {
                    return true;
                }
                else
                {
                    sError = "ข้อมูล Base Unit Quantity ต้องมีค่ามากกว่า 0";
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private bool Valdate_bLot_Number(string Item_Number, string Lot_Number, string connectionString, out string sError, string Item_Client_ID)
        {
            DataTable dt_check = new DataTable();
            sError = string.Empty;
            try
            {
                string sql_display = @" SELECT  Item_Number,bLot_Number,bExpire_date,bAttribute1,bAttribute2,bAttribute3,bAttribute4,UOM,Item_Client_ID
                                       FROM  TMMaster_Validate with(nolock) ";
                sql_display += " Where Item_Number = '" + Item_Number + "' and Item_Client_ID = '" + Item_Client_ID + "'";
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_check);
                if (dt_check.Rows.Count != 0)
                {
                    bool bLot_Number = Convert.ToBoolean(dt_check.Rows[0]["bLot_Number"].ToString());
                    if (bLot_Number)
                    {
                        if (Lot_Number.Trim() != string.Empty)
                        {
                            return true;
                        }
                        else
                        {
                            sError = "Item Number นี้ควรมี Lot Number";
                            return false;
                        }


                    }
                    else // bLot_Number = false
                    {
                        if (Lot_Number.Trim() != string.Empty)
                        {
                            sError = "Item Number นี้ไม่ควรมี Lot Number";
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }

                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private bool Valdate_bAttribute1(string Item_Number, string Attribute1, string connectionString, out string sError, string Item_Client_ID)
        {
            DataTable dt_check = new DataTable();
            sError = string.Empty;
            try
            {
                string sql_display = @" SELECT  Item_Number,bLot_Number,bExpire_date,bAttribute1,bAttribute2,bAttribute3,bAttribute4,UOM,Item_Client_ID
                                       FROM  TMMaster_Validate with(nolock) ";
                sql_display += " Where Item_Number = '" + Item_Number + "' and Item_Client_ID = '" + Item_Client_ID + "'";
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_check);
                if (dt_check.Rows.Count != 0)
                {
                    bool bAttribute1 = Convert.ToBoolean(dt_check.Rows[0]["bAttribute1"].ToString());
                    if (bAttribute1)
                    {
                        //if (Attribute1.Trim() != string.Empty)
                        //{
                        //    return true;
                        //}
                        //else
                        //{
                        //    return false;
                        //}
                        return true;
                    }
                    else // bAttribute1 = false
                    {
                        if (Attribute1.Trim() != string.Empty)
                        {
                            sError = "Item Number ดังกล่าวไม่ควรมีค่า Attribute1";
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }

                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private bool Valdate_bAttribute2(string Item_Number, string Attribute2, string connectionString, out string sError, string Item_Client_ID)
        {
            DataTable dt_check = new DataTable();
            sError = string.Empty;
            try
            {
                string sql_display = @" SELECT  Item_Number,bLot_Number,bExpire_date,bAttribute1,bAttribute2,bAttribute3,bAttribute4,UOM,Item_Client_ID
                                       FROM  TMMaster_Validate with(nolock) ";
                sql_display += " Where Item_Number = '" + Item_Number + "' and Item_Client_ID = '" + Item_Client_ID + "'";
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_check);
                if (dt_check.Rows.Count != 0)
                {
                    bool bAttribute2 = Convert.ToBoolean(dt_check.Rows[0]["bAttribute2"].ToString());
                    if (bAttribute2)
                    {
                        return true;
                    }
                    else // bAttribute1 = false
                    {
                        if (Attribute2.Trim() != string.Empty)
                        {
                            sError = "Item Number ดังกล่าวไม่ควรมีค่า Attribute2";
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }

                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private bool Valdate_bAttribute3(string Item_Number, string Attribute3, string connectionString, out string sError, string Item_Client_ID)
        {
            DataTable dt_check = new DataTable();
            sError = string.Empty;
            try
            {
                string sql_display = @" SELECT  Item_Number,bLot_Number,bExpire_date,bAttribute1,bAttribute2,bAttribute3,bAttribute4,UOM,Item_Client_ID
                                       FROM  TMMaster_Validate with(nolock) ";
                sql_display += " Where Item_Number = '" + Item_Number + "' and Item_Client_ID = '" + Item_Client_ID + "'";
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_check);
                if (dt_check.Rows.Count != 0)
                {
                    bool bAttribute3 = Convert.ToBoolean(dt_check.Rows[0]["bAttribute3"].ToString());
                    if (bAttribute3)
                    {
                        return true;
                    }
                    else // bAttribute1 = false
                    {
                        if (Attribute3.Trim() != string.Empty)
                        {
                            sError = "Item Number ดังกล่าวไม่ควรมีค่า Attribute3";
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }

                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private bool Valdate_bAttribute4(string Item_Number, string Attribute4, string connectionString, out string sError, string Item_Client_ID)
        {
            DataTable dt_check = new DataTable();
            sError = string.Empty;
            try
            {
                string sql_display = @" SELECT  Item_Number,bLot_Number,bExpire_date,bAttribute1,bAttribute2,bAttribute3,bAttribute4,UOM,Item_Client_ID
                                       FROM  TMMaster_Validate with(nolock) ";
                sql_display += " Where Item_Number = '" + Item_Number + "' and Item_Client_ID = '" + Item_Client_ID + "'";
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_check);
                if (dt_check.Rows.Count != 0)
                {
                    bool bAttribute4 = Convert.ToBoolean(dt_check.Rows[0]["bAttribute4"].ToString());
                    if (bAttribute4)
                    {
                        return true;
                    }
                    else // bAttribute1 = false
                    {
                        if (Attribute4.Trim() != string.Empty)
                        {
                            sError = "Item Number ดังกล่าวไม่ควรมีค่า Attribute4";
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }

                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private bool Valdate_Inventory_Status(string Inventory_Status, string connectionString, out string sError)
        {
            DataTable dt_check = new DataTable();
            sError = string.Empty;
            try
            {
                string sql_display = @" SELECT Inventory_Status
                                       FROM  TMInventoryStatus with(nolock) ";
                sql_display += " Where Inventory_Status = '" + Inventory_Status + "'";
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_check);
                if (dt_check.Rows.Count != 0)
                {

                    return true;
                }
                else
                {
                    sError = "Inventory Status : " + Inventory_Status+ " ดังกล่าวไม่มีในตาราง Master";
                    return false;
                }
            }
            catch (Exception ex)
            {
                sError = ex.Message.ToString();
                return false;
            }
        }
        private bool Valdate_bExpire_date(string Item_Number, string Lot_Number, string Expiration_Date, string connectionString, out string sError, string Item_Client_ID)
        {
            DataTable dt_check = new DataTable();
            sError = string.Empty;
            try
            {
                string sql_display = @" SELECT  Item_Number,bLot_Number,bExpire_date,bAttribute1,bAttribute2,bAttribute3,bAttribute4,UOM,Item_Client_ID
                                       FROM  TMMaster_Validate with(nolock) ";
                sql_display += " Where Item_Number = '" + Item_Number + "' and Item_Client_ID = '" + Item_Client_ID + "'";
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_check);
                if (dt_check.Rows.Count != 0)
                {
                    bool bLot_Number = Convert.ToBoolean(dt_check.Rows[0]["bLot_Number"].ToString());
                    if (bLot_Number)
                    {
                        if (Expiration_Date != string.Empty)
                        {
                            return true;
                        }
                        else
                        {
                            sError = "Item Number นี้ต้องมีวันที่ Expiration Date";
                            return false;
                        }
                    }
                    else // bLot_Number = false
                    {
                        if (Expiration_Date.Trim() != string.Empty)
                        {
                            sError = "Item Number นี้ต้องไม่มีวันที่ Expiration Date";
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }

                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private bool validate_Date(string sDate, out string error)
        {
            error = string.Empty;
            try
            {
                string[] sArray = sDate.Split('/');
                if (sArray.Length == 3)
                {
                    if (sArray[0].Length != 4)
                    {
                        error = "Format date not correct";
                        return false;
                    }
                    else
                    {
                        if (Convert.ToInt32(sArray[1]) > 12)
                        {
                            error = "Format date not correct";
                            return false;
                        }
                        else
                        {
                            return true;
                        }
                    }
                }
                else
                {
                    error = "Format date not correct";
                    return false;
                }
            
            }
            catch (Exception ex)
            {
                error = ex.Message.ToString();
                return false;
            }
        }
        private bool Valdate_MaxExpire_date(string Expiration_Date, out string error)
        {
            error = string.Empty;
            int MaxExpire = 20991231;
            int MaxExpire_sys = 0;
            bool bResult = false;
            try
            {
                //bool bvalidate_Date = validate_Date(Expiration_Date, out error);
                if (validate_Date(Expiration_Date, out error))
                {
                    if (Expiration_Date.Trim() != string.Empty)
                    {
                        MaxExpire_sys = Convert.ToInt32(Expiration_Date.Substring(0, 4) + Expiration_Date.Substring(5, 2) + Expiration_Date.Substring(8, 2));
                        if (MaxExpire_sys > MaxExpire)
                        {
                            error = "Expiration Date มีค่ามากกว่าวันที่ 2099/12/31";
                            bResult = false;
                        }
                        else
                        {
                            bResult = true;
                        }
                    }
                }
                else
                {
                    bResult = false;
                }
                return bResult;
            }
            catch (Exception ex)
            {
                error = ex.ToString() + "Format not correct, Please check format";
                return false;
            }
        }

        #endregion
        private void Button_Validate_Click(object sender, EventArgs e)
        {
            ETMItemBlance lst = new ETMItemBlance();
            Cursor.Current = Cursors.WaitCursor;

            string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
            try
            {
                DataTable dt_max = new DataTable();
                string sql_Max = @" SELECT max(nRound) nRound FROM TMItemBlance with(nolock)";
                SqlDataAdapter da = new SqlDataAdapter(sql_Max, connectionString);
                da.Fill(dt_max);

                dt = new DataTable();
                string sql_display = @" SELECT Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number
                                       ,Supplier_Lot_Number,Received_Date,Manufactured_Date,Expiration_Date,Base_Unit_Qty
                                       ,Base_Unit_UOM,Inventory_Status,Attribute_1,Attribute_2,Attribute_3,Attribute_4
                                        FROM  TMItemBlance with(nolock) ";
                sql_display += " Where nRound = " + dt_max.Rows[0]["nRound"].ToString() + "";
                da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt);
                //dataGridView_Display.Visible = true;
                //dataGridView_Display.DataSource = dt;
                bool bCheck_Error = false;
                bool bDelete = Delete_TMItemBlance_Valid();
                int ProgressMinimum = 0;
                ProgressMaximum = dt.Rows.Count;
                int ProgressValue = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    bool bError = false;

                    lst = new ETMItemBlance();
                    lst.Warehouse_ID = dt.Rows[i]["Warehouse_ID"].ToString();
                    lst.Location_Code = dt.Rows[i]["Location_Code"].ToString();
                    lst.LPN = dt.Rows[i]["LPN"].ToString();
                    lst.Item_Client_ID = dt.Rows[i]["Item_Client_ID"].ToString();
                    lst.Item_Number = dt.Rows[i]["Item_Number"].ToString();
                    lst.Lot_Number = dt.Rows[i]["Lot_Number"].ToString();
                    lst.Supplier_Lot_Number = dt.Rows[i]["Supplier_Lot_Number"].ToString();
                    lst.Received_Date = dt.Rows[i]["Received_Date"].ToString();
                    lst.Manufactured_Date = dt.Rows[i]["Manufactured_Date"].ToString();
                    lst.Expiration_Date = dt.Rows[i]["Expiration_Date"].ToString();
                    lst.Base_Unit_Qty = dt.Rows[i]["Base_Unit_Qty"].ToString();
                    lst.Base_Unit_UOM = dt.Rows[i]["Base_Unit_UOM"].ToString();
                    lst.Inventory_Status = dt.Rows[i]["Inventory_Status"].ToString();
                    lst.Attribute_1 = dt.Rows[i]["Attribute_1"].ToString();
                    lst.Attribute_2 = dt.Rows[i]["Attribute_2"].ToString();
                    lst.Attribute_3 = dt.Rows[i]["Attribute_3"].ToString();
                    lst.Attribute_4 = dt.Rows[i]["Attribute_4"].ToString();

                    lst.sAttribute1 = string.Empty;
                    lst.sAttribute2 = string.Empty;
                    lst.sAttribute3 = string.Empty;
                    lst.sAttribute4 = string.Empty;
                    lst.sBase_Qty = string.Empty;
                    lst.sCheck_Expire_date = string.Empty;
                    lst.sExpire_date = string.Empty;
                    lst.sInventory_Status = string.Empty;
                    lst.sItem_Number = string.Empty;
                    lst.sLot_Number = string.Empty;
                    lst.sUOM = string.Empty;
                    lst.sCheckManufactured_Date = string.Empty;
                    lst.sCheckReceived_Date = string.Empty;

                    string sError = string.Empty;
                    if (Valdate_Item_Number(lst.Item_Number, connectionString, dt.Rows[i]["Item_Client_ID"].ToString())) // Found Item_Number
                    {
                        lst.sItem_Number = "";

                        if (lst.Received_Date.Trim() != string.Empty)
                        {
                            if (!validate_Date(lst.Received_Date.Trim(), out sError))
                            {
                                bError = true;
                                lst.sCheckReceived_Date = sError;
                            }
                        }

                        if (lst.Manufactured_Date.Trim() != string.Empty)
                        {
                            if (!validate_Date(lst.Manufactured_Date.Trim(), out sError))
                            {
                                bError = true;
                                lst.sCheckManufactured_Date = sError;
                            }
                        }

                        if (!Valdate_UOM(lst.Item_Number, lst.Base_Unit_UOM, connectionString, out sError, dt.Rows[i]["Item_Client_ID"].ToString()))
                        {
                            bError = true;
                            lst.sUOM = sError;
                        }
                        else
                        {
                            lst.sUOM = "";
                        }

                        if (!Valdate_bLot_Number(lst.Item_Number, lst.Lot_Number, connectionString, out sError, dt.Rows[i]["Item_Client_ID"].ToString()))
                        {
                            bError = true;
                            lst.sLot_Number = sError;
                        }
                        else
                        {
                            lst.sLot_Number = "";
                        }

                        if (!Valdate_bExpire_date(lst.Item_Number, lst.Lot_Number, lst.Expiration_Date, connectionString, out sError, dt.Rows[i]["Item_Client_ID"].ToString()))
                        {
                            bError = true;
                            lst.sExpire_date = sError;
                        }
                        else
                        {
                            lst.sExpire_date = "";
                        }

                        if (!Valdate_Base_Unit_Qty(lst.Base_Unit_Qty, out sError))
                        {
                            bError = true;
                            lst.sBase_Qty = sError;
                        }
                        else
                        {
                            lst.sBase_Qty = "";
                        }

                        if (!Valdate_bAttribute1(lst.Item_Number, lst.Attribute_1, connectionString, out sError, dt.Rows[i]["Item_Client_ID"].ToString()))
                        {
                            bError = true;
                            lst.sAttribute1 = sError;
                        }
                        else
                        {
                            lst.sAttribute1 = "";
                        }

                        if (!Valdate_bAttribute2(lst.Item_Number, lst.Attribute_2, connectionString, out sError, dt.Rows[i]["Item_Client_ID"].ToString()))
                        {
                            bError = true;
                            lst.sAttribute2 = sError;
                        }
                        else
                        {
                            lst.sAttribute2 = "";
                        }

                        if (!Valdate_bAttribute3(lst.Item_Number, lst.Attribute_3, connectionString, out sError, dt.Rows[i]["Item_Client_ID"].ToString()))
                        {
                            bError = true;
                            lst.sAttribute3 = sError;
                        }
                        else
                        {
                            lst.sAttribute3 = "";
                        }

                        if (!Valdate_bAttribute4(lst.Item_Number, lst.Attribute_4, connectionString, out sError, dt.Rows[i]["Item_Client_ID"].ToString()))
                        {
                            bError = true;
                            lst.sAttribute4 = sError;
                        }
                        else
                        {
                            lst.sAttribute4 = "";
                        }

                        if (!Valdate_Inventory_Status(lst.Inventory_Status, connectionString, out sError))
                        {
                            bError = true;
                            lst.sInventory_Status = sError;
                        }
                        else
                        {
                            lst.sInventory_Status = "";
                        }

                        if (lst.Expiration_Date != string.Empty)
                        {
                            if (!Valdate_MaxExpire_date(lst.Expiration_Date, out sError))
                            {
                                bError = true;
                                lst.sCheck_Expire_date = sError;
                            }
                            else
                            {
                                lst.sCheck_Expire_date = "";
                            }
                        }
                    }
                    else // not found Item_Number
                    {
                        bError = true;
                        lst.sItem_Number = "ไม่เจอ Item no. ในระบบ";
                        lst.sLot_Number = "";
                        lst.sExpire_date = "";
                        lst.sAttribute1 = "";
                        lst.sAttribute2 = "";
                        lst.sAttribute3 = "";
                        lst.sAttribute4 = "";
                        lst.sInventory_Status = "";
                        lst.sUOM = "";
                        lst.sBase_Qty = "";
                        lst.sCheck_Expire_date = "";
                    }

                    if (bError) // Found Error Insert TMItemBlance_Valid
                    {
                        bCheck_Error = true;
                        Save_TMItemBlance_Valid(lst);
                    }
                    else // Not Found Error Insert TMItemBlance_Pass
                    {
                        Save_TMItemBlance_Pass(lst);
                    }
                }

                if (bCheck_Error)
                {
                    MessageBox.Show("เจอข้อมูลบางรายการไม่ถูกต้องตอนตรวจสอบข้อมูล");
                    Show_Valid();
                }
                else
                {
                    MessageBox.Show("ข้อมุลไม่มีข้อผิดพลาด");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void Button_Complete_Click(object sender, EventArgs e)
        {
            Show_Pass();
        }

        private void Button_ExportExpireDate_Click(object sender, EventArgs e)
        {
            if (textBox_Path.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please select path before");
                return;
            }

            if (textBox_Name.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Please input file name before");
                return;
            }

            Cursor.Current = Cursors.WaitCursor;
            string connectionString = System.Configuration.ConfigurationSettings.AppSettings["Constr"].ToString();
            try
            {

                DataTable dt_ex = new DataTable();
                string sql_display = @"
Select Warehouse_ID, Location_Code, LPN, Item_Client_ID, Item_Number, Lot_Number, Supplier_Lot_Number, Received_Date, Manufactured_Date, Expiration_Date, Base_Unit_Qty, Base_Unit_UOM, Inventory_Status, 
                                                    Attribute_1, Attribute_2, Attribute_3, Attribute_4 
from
(
SELECT        Warehouse_ID, Location_Code, LPN, Item_Client_ID, Item_Number, Lot_Number, Supplier_Lot_Number, Received_Date, Manufactured_Date, Expiration_Date, Base_Unit_Qty, Base_Unit_UOM, Inventory_Status, 
                                                    Attribute_1, Attribute_2, Attribute_3, Attribute_4, nID
                          FROM            dbo.TMHeader WITH (nolock)

						  UNION 

Select Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number
,Supplier_Lot_Number,Received_Date,Manufactured_Date,Expiration_Date, CAST(Base_Unit_Qty AS nvarchar) AS Base_Unit_Qty
,Base_Unit_UOM,Inventory_Status,Attribute_1,Attribute_2,Attribute_3,Attribute_4,99 as nID
from
(
SELECT  Warehouse_ID,Location_Code,LPN,Item_Client_ID,Item_Number,Lot_Number
,Supplier_Lot_Number,Received_Date,Manufactured_Date,Expiration_Date,Base_Unit_Qty
,Base_Unit_UOM,Inventory_Status,Attribute_1,Attribute_2,Attribute_3,Attribute_4,Item_Number+Lot_Number+Expiration_Date  Expiration_Date2
  FROM  TMItemBlance_Pass WITH (nolock)
  Where Item_Number+Lot_Number in (Select ItemLot
from 
(
SELECT distinct Item_Number, Lot_Number ,Item_Number+Lot_Number as ItemLot,   Expiration_Date
FROM   dbo.TMItemBlance_Pass AS TMItemBlance_Pass_1 WITH (nolock)
) as TMItemBlance_Pass
Where Expiration_Date <> ''
group by Item_Number, Lot_Number ,ItemLot
having count(Expiration_Date)  > 1)
 ) as TMItemBlance_Pass_ExpireDate
 Where Expiration_Date2 not in (Select Expiration_Date2
from
(
Select Item_Number,Lot_Number,Expiration_Date ,Item_Number+Lot_Number+Expiration_Date Expiration_Date2
from
(
SELECT Item_Number, Lot_Number, MIN(Expiration_Date) AS Expiration_Date 
                               FROM            dbo.TMItemBlance_Pass AS TMItemBlance_Pass_1 WITH (nolock)
							   Where Item_Number +Lot_Number in (Select ItemLot
from 
(
SELECT distinct Item_Number, Lot_Number ,Item_Number+Lot_Number as ItemLot,   Expiration_Date
FROM   dbo.TMItemBlance_Pass AS TMItemBlance_Pass_1 WITH (nolock)
) as TMItemBlance_Pass
Where Expiration_Date <> ''
group by Item_Number, Lot_Number ,ItemLot
having count(Expiration_Date)  > 1)
GROUP BY Item_Number, Lot_Number
) as TMItemBlance_Pass
) as TMItemBlance_Pass)
 
 ) as TMItemBlance_Pass
 order by nID ,Item_Number, Lot_Number ,Expiration_Date"; // 
                SqlDataAdapter da = new SqlDataAdapter(sql_display, connectionString);
                da.Fill(dt_ex);
                if (dt_ex.Rows.Count > 2)
                {
                    Export_Excel(dt_ex);
                }
                else
                {
                    MessageBox.Show("Not found data");
                    return;
                }

                //textBox_Name.Text = string.Empty;
                textBox_Path.Text = string.Empty;
                dt = new DataTable();
                dt = AddColumnDatatable();
                dataGridView_Display.Visible = true;
                dataGridView_Display.DataSource = dt_ex;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                return;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

    }
}
