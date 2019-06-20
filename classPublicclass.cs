using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Security.AccessControl;
using System.Security.Principal;
using iTextSharp.text.pdf;
using System.Diagnostics;

namespace DataExplorer2
{
    class ClassPublicclass
    {
        public static void Numbersonly(object sender, KeyPressEventArgs e)
        {
            char keyChar;
            keyChar = e.KeyChar;

            if (!char.IsDigit(keyChar) // 0 - 9
                &&
                keyChar != 8 // backspace
                &&
                keyChar != 13 // enter
                &&
                keyChar != '.'
                &&
                keyChar != 45 //  dash/minus
            )
                e.Handled = true;
        }

        public static string GetTemporaryDirectory(string dirc)
        {
            var tempDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DataExplorer";
            try
            {
                var pathDocuments = Path.Combine(tempDirectory, dirc);
                Directory.CreateDirectory(pathDocuments);
            }
            catch (Exception)
            {
            }

            return tempDirectory;
        }


        public static void SelectfromtableLBL(Label Targetlbl, string TableName, string FieldName, string condition)
        {
            try
            {
                //fill a referenced combobox with a referenced field from a referenced table
                var sbConnCOMBO = new StringBuilder();
                sbConnCOMBO.Append(ClassDatabaseConnection.cnn1);
                sbConnCOMBO.Append(";Extended Properties=");
                sbConnCOMBO.Append(Convert.ToChar(34));
                sbConnCOMBO.Append(Convert.ToChar(34));

                var cnExcelCOMBO = new OleDbConnection(sbConnCOMBO.ToString());
                cnExcelCOMBO.Open();

                var sbSQLCOMBO = new StringBuilder();
                if (TableName == "INFORMATION_SCHEMA.COLUMNS IC") sbSQLCOMBO.Append("Select [" + FieldName + "] FROM  " + TableName + " " + condition);

                sbSQLCOMBO.Append("Select [" + FieldName + "] FROM [" + TableName + "]" + condition);
                // classPublicclass.showMessage("Select [" + FieldName + "] FROM [" + TableName + "]" + condition);
                var cmdExcelCOMBO = new OleDbCommand(sbSQLCOMBO.ToString(), cnExcelCOMBO);
                var drExcelCOMBO = cmdExcelCOMBO.ExecuteReader();

                while (drExcelCOMBO.Read()) Targetlbl.Text = drExcelCOMBO[FieldName].ToString();

                cnExcelCOMBO.Close();
            }
            catch (Exception)
            {
            }
        }



        public static void GenerateFile(string content)
        {
            var name = "mytestFile";
            var extension = "txt";
            var folder = "DE_Reports";
            // path = classPublicclass.GetTemporaryDirectory("DE_Reports");
           var path = GetTemporaryDirectory(folder);
            var docname = name;
            // string extension = "txt";
            try
            {
                // create a writer and open the file
                File.Delete(string.Format("{0}\\{1}.{2}", path + "\\" + folder + "\\", docname, extension));
                GrantAccess(@path + "\\" + folder + "\\" + docname + "." + extension);
                File.Delete(@path + "\\" + folder + "\\" + docname + "." + extension);
                //TextWriter tw = new StreamWriter(docname + ".doc");
                TextWriter tw =
                    new StreamWriter(string.Format("{0}\\{1}.{2}", path + "\\" + folder + "\\", docname, extension));
                // write a line of text to the file
                tw.WriteLine(content);
                tw.Close();

                try
                {
                    Process.Start(@"" + string.Format("{0}\\{1}.{2}", path + "\\" + folder + "\\",
                                                         docname, extension));
                }
                catch (Exception n)
                {
                    MessageBox.Show("Document may be open please close it and try again " + n.Message);
                }
            }
            catch (Exception)
            {
            }
        }


        public static void Filllist(ListView lsvMain, string sql) // POPULATE MAIN LISTVIEW
        {
            
            try
            {
                // Cursor.Current = Cursors.WaitCursor;
                var sbConn = new StringBuilder();
                sbConn.Append(ClassDatabaseConnection.cnn1); // CONNECT TO DATABASE
               
                sbConn.Append(";Extended Properties=");
                sbConn.Append(Convert.ToChar(34));
                sbConn.Append(Convert.ToChar(34));
                //
                //
                // Open the database and query the data.
                //
                var cnExcel = new OleDbConnection(sbConn.ToString());
                cnExcel.Open();
                var sbSQL = new StringBuilder();
                sbSQL.Append(sql);
                var cmdExcel = new OleDbCommand(sbSQL.ToString(), cnExcel);
                var drExcel = cmdExcel.ExecuteReader();

                if (drExcel.FieldCount > 0)
                {
                    lsvMain.Items.Clear();
                    lsvMain.Columns.Clear();
                    for (var i = 0; i < drExcel.FieldCount; i++)
                        if (i == 0)
                        {
                            lsvMain.Columns.Add(drExcel.GetName(0), 0, HorizontalAlignment.Left);
                            ClassMainclass.focussearchitem = drExcel.GetName(0).ToString();
                        }
                        else
                        {
                            if (
                                drExcel.GetName(i).ToString() == "image" ||
                                drExcel.GetName(i).ToString() == "Image" ||
                                drExcel.GetName(i).ToString() == "photo" ||
                                drExcel.GetName(i).ToString() == "Photo" ||
                                drExcel.GetName(i).ToString() == "Password" ||
                                drExcel.GetName(i).ToString() == "password" ||
                                drExcel.GetName(i).ToString() == "Logo" ||                              
                                drExcel.GetName(i).ToString() == "logo")
                            {
                            }
                            else
                            {

                                // ALIGHN QUANTITATIVE DATA TO THE RIGHT
                                   if (drExcel.GetName(i).ToString() == "Miles Per Gallon" ||
                                    drExcel.GetName(i).ToString() == "Litres Per 100 km" ||
                                    drExcel.GetName(i).ToString() == "Vatable Total" ||
                                    drExcel.GetName(i).ToString() == "Subtotal" ||
                                    drExcel.GetName(i).ToString() == "Tax" ||
                                    drExcel.GetName(i).ToString() == "Total" ||
                                    drExcel.GetName(i).ToString() == "Debit Amount" ||
                                    drExcel.GetName(i).ToString() == "Total VAT" ||
                                    drExcel.GetName(i).ToString() == "Vatable Total" ||
                                    drExcel.GetName(i).ToString() == "Paid Amount" ||
                                    drExcel.GetName(i).ToString() == "Balance" ||
                                    drExcel.GetName(i).ToString() == "Consumed Qty" ||
                                    drExcel.GetName(i).ToString() == "Fine Amount" ||
                                    drExcel.GetName(i).ToString() == "Qty" ||
                                    drExcel.GetName(i).ToString() == "Initial Qty" ||
                                    drExcel.GetName(i).ToString() == "VAT AMOUNT" ||
                                    drExcel.GetName(i).ToString() == "Debit Amount" ||
                                    drExcel.GetName(i).ToString() == "Netchange Debit" ||
                                    drExcel.GetName(i).ToString() == "Netchange Credit" ||
                                    drExcel.GetName(i).ToString() == "Balance Debit" ||
                                    drExcel.GetName(i).ToString() == "Balance Credit" ||
                                    drExcel.GetName(i).ToString() == "Credit Amount" ||
                                    drExcel.GetName(i).ToString() == "VAT%" ||
                                    drExcel.GetName(i).ToString() == "Total Cost" ||
                                    drExcel.GetName(i).ToString() == "VAT Amount" ||
                                    drExcel.GetName(i).ToString() == "Unit Cost" ||
                                    drExcel.GetName(i).ToString() == "Balance Credit" ||
                                    drExcel.GetName(i).ToString() == "Balance Debit" ||
                                    drExcel.GetName(i).ToString() == "Net Change Credit" ||
                                    drExcel.GetName(i).ToString() == "Net Change Debit" ||
                                    drExcel.GetName(i).ToString() == "Net Amount" ||
                                    drExcel.GetName(i).ToString() == "Receipt Amount" ||
                                    drExcel.GetName(i).ToString() == "Total Deductions" ||
                                    drExcel.GetName(i).ToString() == "Pension" ||
                                    drExcel.GetName(i).ToString() == "NHIF" ||
                                    drExcel.GetName(i).ToString() == "NSSF" ||
                                    drExcel.GetName(i).ToString() == "PAYE" || drExcel.GetName(i).ToString() == "MPR" ||
                                    drExcel.GetName(i).ToString() == "Tax Charged" ||
                                    drExcel.GetName(i).ToString() == "Net Pay" ||
                                    drExcel.GetName(i).ToString() == "Taxable Pay" ||
                                    drExcel.GetName(i).ToString() == "Gross Pay" ||
                                    drExcel.GetName(i).ToString() == "Basic Pay" ||
                                    drExcel.GetName(i).ToString() == "Amount" ||
                                    drExcel.GetName(i).ToString() == "Cost" ||
                                    drExcel.GetName(i).ToString() == "Price" ||
                                    drExcel.GetName(i).ToString() == "Day Weighting")
                                    lsvMain.Columns.Add(drExcel.GetName(i).ToString(), 80, HorizontalAlignment.Right);
                                else
                                    lsvMain.Columns.Add(drExcel.GetName(i).ToString().Replace("_", " "), 80,
                                        HorizontalAlignment.Left);
                            }
                        }

                    var lv = new ListViewItem();
                    //
                    while (drExcel.Read())
                    {
                        lv = lsvMain.Items.Add(drExcel[drExcel.GetName(0)].ToString().Replace('_', ' '));

                       
                        for (var h = 1; h < drExcel.FieldCount; h++)
                            if (
                                // CLEAN SENSITIVE AND IRRELEVANT  DATA 
                                drExcel.GetName(h).ToString() == "image" ||
                                drExcel.GetName(h).ToString() == "photo" ||
                                drExcel.GetName(h).ToString() == "Image" ||
                                drExcel.GetName(h).ToString() == "Photo" ||
                                drExcel.GetName(h).ToString() == "Logo" ||                      
                                drExcel.GetName(h).ToString() == "Password")
                            {
                            }
                            else
                            {
                                // ALIGHN QUANTITATIVE DATA TO THE RIGHT
                                var checker = "";
                                if (drExcel.GetName(h).ToString() == "Miles Per Gallon" ||
                                    drExcel.GetName(h).ToString() == "Litres Per 100 km" ||
                                    drExcel.GetName(h).ToString() == "Fine Amount" ||
                                    drExcel.GetName(h).ToString() == "Vatable Total" ||
                                    drExcel.GetName(h).ToString() == "Subtotal" ||
                                    drExcel.GetName(h).ToString() == "Tax" ||
                                    drExcel.GetName(h).ToString() == "Total" ||
                                    drExcel.GetName(h).ToString() == "Debit Amount" ||
                                    drExcel.GetName(h).ToString() == "Total VAT" ||
                                    drExcel.GetName(h).ToString() == "Vatable Total" ||
                                    drExcel.GetName(h).ToString() == "Paid Amount" ||
                                    drExcel.GetName(h).ToString() == "Balance" ||
                                    drExcel.GetName(h).ToString() == "Consumed Qty" ||
                                    drExcel.GetName(h).ToString() == "Initial Qty" ||
                                    drExcel.GetName(h).ToString() == "Qty" ||
                                    drExcel.GetName(h).ToString() == "VAT AMOUNT" ||
                                    drExcel.GetName(h).ToString() == "Debit Amount" ||
                                    drExcel.GetName(h).ToString() == "Netchange Debit" ||
                                    drExcel.GetName(h).ToString() == "Netchange Credit" ||
                                    drExcel.GetName(h).ToString() == "Balance Debit" ||
                                    drExcel.GetName(h).ToString() == "Balance Credit" ||
                                    drExcel.GetName(h).ToString() == "Credit Amount" ||
                                    drExcel.GetName(h).ToString() == "VAT%" ||
                                    drExcel.GetName(h).ToString() == "Total Cost" ||
                                    drExcel.GetName(h).ToString() == "VAT Amount" ||
                                    drExcel.GetName(h).ToString() == "Unit Cost" ||
                                    drExcel.GetName(h).ToString() == "Balance Credit" ||
                                    drExcel.GetName(h).ToString() == "Balance Debit" ||
                                    drExcel.GetName(h).ToString() == "Net Change Credit" ||
                                    drExcel.GetName(h).ToString() == "Net Change Debit" ||
                                    drExcel.GetName(h).ToString() == "Net Amount" ||
                                    drExcel.GetName(h).ToString() == "Receipt Amount" ||
                                    drExcel.GetName(h).ToString() == "Total Deductions" ||
                                    drExcel.GetName(h).ToString() == "Pension" ||
                                    drExcel.GetName(h).ToString() == "NHIF" ||
                                    drExcel.GetName(h).ToString() == "NSSF" ||
                                    drExcel.GetName(h).ToString() == "PAYE" || drExcel.GetName(h).ToString() == "MPR" ||
                                    drExcel.GetName(h).ToString() == "Tax Charged" ||
                                    drExcel.GetName(h).ToString() == "Net Pay" ||
                                    drExcel.GetName(h).ToString() == "Taxable Pay" ||
                                    drExcel.GetName(h).ToString() == "Gross Pay" ||
                                    drExcel.GetName(h).ToString() == "Basic Pay" ||
                                    drExcel.GetName(h).ToString() == "Amount" ||
                                    drExcel.GetName(h).ToString() == "Cost" ||
                                    drExcel.GetName(h).ToString() == "Price" ||
                                    drExcel.GetName(h).ToString() == "Day Weighting")
                                {
                                    if (drExcel[drExcel.GetName(h)].ToString() == "")
                                        checker = "0";
                                    else
                                        checker = drExcel[drExcel.GetName(h)].ToString();

                                    lv.SubItems.Add(string.Format("{0:0.00}", double.Parse(checker)));
                                }
                                else
                                {
                                    lv.SubItems.Add(drExcel[drExcel.GetName(h)].ToString());
                                }
                            }
                    }
                }

                for (var i = 1; i < lsvMain.Columns.Count; i++)
                    lsvMain.Columns[i].Width = -2;
                cnExcel.Close();
                // Cursor.Current = Cursors.Arrow;
            }
            catch (Exception)
            {
               
            }

            
        }

        public static void Selectfromtablecountlbl(Label Targetlbl, string TableName, string FieldName,
          string condition)
        {
            try
            {
                //fill a referenced combobox with a referenced field from a referenced table
                var sbConnCOMBO = new StringBuilder();
                sbConnCOMBO.Append(ClassDatabaseConnection.cnn1);
                sbConnCOMBO.Append(";Extended Properties=");
                sbConnCOMBO.Append(Convert.ToChar(34));
                sbConnCOMBO.Append(Convert.ToChar(34));

                var cnExcelCOMBO = new OleDbConnection(sbConnCOMBO.ToString());
                cnExcelCOMBO.Open();

                var sbSQLCOMBO = new StringBuilder();
                sbSQLCOMBO.Append("Select [" + FieldName + "] FROM [" + TableName + "]" + condition);

                var cmdExcelCOMBO = new OleDbCommand(sbSQLCOMBO.ToString(), cnExcelCOMBO);
                var drExcelCOMBO = cmdExcelCOMBO.ExecuteReader();

                var cnt = 0;

                while (drExcelCOMBO.Read())
                    // Targetlbl.Text = drExcelCOMBO[FieldName].ToString();
                    //Targetlbl.Text = drExcelCOMBO.RecordsAffected.ToString();
                    cnt++;

                Targetlbl.Text = cnt.ToString();
                cnExcelCOMBO.Close();
            }
            catch (Exception)
            {
            }
        }


        public static void Filllistnoid(string tablename)
        {
            //  colorListViewHeader(lsvMain);
            var lsvMain = new ListView();

            try
            {
                // Cursor.Current = Cursors.WaitCursor;
                var sbConn = new StringBuilder();
                sbConn.Append(ClassDatabaseConnection.cnn1);
                sbConn.Append(";Extended Properties=");
                sbConn.Append(Convert.ToChar(34));
                sbConn.Append(Convert.ToChar(34));
                //
                //
                // Open the database and query the data.
                //
                var cnExcel = new OleDbConnection(sbConn.ToString());
                cnExcel.Open();
                var sbSQL = new StringBuilder();
                sbSQL.Append("SELECT * FROM '" + tablename + "'");
                var cmdExcel = new OleDbCommand(sbSQL.ToString(), cnExcel);
                var drExcel = cmdExcel.ExecuteReader();

                if (drExcel.FieldCount > 0)
                {
                    lsvMain.Items.Clear();
                    lsvMain.Columns.Clear();
                    for (var i = 0; i < drExcel.FieldCount; i++)
                        if (i == 0)
                        {
                            lsvMain.Columns.Add(drExcel.GetName(0), 0, HorizontalAlignment.Left);
                            ClassMainclass.focussearchitem = drExcel.GetName(0).ToString();
                        }
                        else
                        {
                            if (
                                drExcel.GetName(i).ToString() == "image" ||
                                drExcel.GetName(i).ToString() == "Photo" ||
                                drExcel.GetName(i).ToString() == "Logo" ||
                                drExcel.GetName(i).ToString() == "ID")
                            {
                            }
                            else
                            {
                                //if (drExcel.GetName(i).ToString() == "ISBN")
                                // {
                                //   lsvMain.Columns.Add("ISBN", 80, HorizontalAlignment.Left);
                                // }
                                // if (drExcel.GetName(i).ToString() == "Net Change Credit" || drExcel.GetName(i).ToString() == "Net Change Debit" || drExcel.GetName(i).ToString() == "Debit Amount" || drExcel.GetName(i).ToString() == "Netchange Debit" || drExcel.GetName(i).ToString() == "Netchange Credit" || drExcel.GetName(i).ToString() == "Balance Debit" || drExcel.GetName(i).ToString() == "Balance Credit" || drExcel.GetName(i).ToString() == "Credit Amount" || drExcel.GetName(i).ToString() == "VAT%" || drExcel.GetName(i).ToString() == "Total Cost" || drExcel.GetName(i).ToString() == "VAT Amount" || drExcel.GetName(i).ToString() == "Unit Cost" || drExcel.GetName(i).ToString() == "Quantity" || drExcel.GetName(i).ToString() == "Balance" || drExcel.GetName(i).ToString() == "Total Deductions" || drExcel.GetName(i).ToString() == "Pension" || drExcel.GetName(i).ToString() == "NHIF" || drExcel.GetName(i).ToString() == "NSSF" || drExcel.GetName(i).ToString() == "PAYE" || drExcel.GetName(i).ToString() == "MPR" || drExcel.GetName(i).ToString() == "Tax Charged" || drExcel.GetName(i).ToString() == "Net Pay" || drExcel.GetName(i).ToString() == "Taxable Pay" || drExcel.GetName(i).ToString() == "Gross Pay" || drExcel.GetName(i).ToString() == "Basic Pay" || drExcel.GetName(i).ToString() == "Amount" || drExcel.GetName(i).ToString() == "Cost" || drExcel.GetName(i).ToString() == "Price" || drExcel.GetName(i).ToString() == "Day Weighting")
                                if (drExcel.GetName(i).ToString() == "Miles Per Gallon" ||
                                    drExcel.GetName(i).ToString() == "Litres Per 100 km" ||
                                    drExcel.GetName(i).ToString() == "Vatable Total" ||
                                    drExcel.GetName(i).ToString() == "Subtotal" ||
                                    drExcel.GetName(i).ToString() == "Tax" ||
                                    drExcel.GetName(i).ToString() == "Total" ||
                                    drExcel.GetName(i).ToString() == "Debit Amount" ||
                                    drExcel.GetName(i).ToString() == "Total VAT" ||
                                    drExcel.GetName(i).ToString() == "Vatable Total" ||
                                    drExcel.GetName(i).ToString() == "Paid Amount" ||
                                    drExcel.GetName(i).ToString() == "Balance" ||
                                    drExcel.GetName(i).ToString() == "Consumed Qty" ||
                                    drExcel.GetName(i).ToString() == "Fine Amount" ||
                                    drExcel.GetName(i).ToString() == "Qty" ||
                                    drExcel.GetName(i).ToString() == "Initial Qty" ||
                                    drExcel.GetName(i).ToString() == "VAT AMOUNT" ||
                                    drExcel.GetName(i).ToString() == "Debit Amount" ||
                                    drExcel.GetName(i).ToString() == "Netchange Debit" ||
                                    drExcel.GetName(i).ToString() == "Netchange Credit" ||
                                    drExcel.GetName(i).ToString() == "Balance Debit" ||
                                    drExcel.GetName(i).ToString() == "Balance Credit" ||
                                    drExcel.GetName(i).ToString() == "Credit Amount" ||
                                    drExcel.GetName(i).ToString() == "VAT%" ||
                                    drExcel.GetName(i).ToString() == "Total Cost" ||
                                    drExcel.GetName(i).ToString() == "VAT Amount" ||
                                    drExcel.GetName(i).ToString() == "Unit Cost" ||
                                    drExcel.GetName(i).ToString() == "Balance Credit" ||
                                    drExcel.GetName(i).ToString() == "Balance Debit" ||
                                    drExcel.GetName(i).ToString() == "Net Change Credit" ||
                                    drExcel.GetName(i).ToString() == "Net Change Debit" ||
                                    drExcel.GetName(i).ToString() == "Net Amount" ||
                                    drExcel.GetName(i).ToString() == "Receipt Amount" ||
                                    drExcel.GetName(i).ToString() == "Total Deductions" ||
                                    drExcel.GetName(i).ToString() == "Pension" ||
                                    drExcel.GetName(i).ToString() == "NHIF" ||
                                    drExcel.GetName(i).ToString() == "NSSF" ||
                                    drExcel.GetName(i).ToString() == "PAYE" || drExcel.GetName(i).ToString() == "MPR" ||
                                    drExcel.GetName(i).ToString() == "Tax Charged" ||
                                    drExcel.GetName(i).ToString() == "Net Pay" ||
                                    drExcel.GetName(i).ToString() == "Taxable Pay" ||
                                    drExcel.GetName(i).ToString() == "Gross Pay" ||
                                    drExcel.GetName(i).ToString() == "Basic Pay" ||
                                    drExcel.GetName(i).ToString() == "Amount" ||
                                    drExcel.GetName(i).ToString() == "Cost" ||
                                    drExcel.GetName(i).ToString() == "Price" ||
                                    drExcel.GetName(i).ToString() == "Day Weighting")
                                    lsvMain.Columns.Add(drExcel.GetName(i).ToString(), 80, HorizontalAlignment.Right);
                                else
                                    lsvMain.Columns.Add(drExcel.GetName(i).ToString().Replace("_", " "), 80,
                                        HorizontalAlignment.Left);
                            }
                        }

                    var lv = new ListViewItem();
                    //
                    while (drExcel.Read())
                    {
                        lv = lsvMain.Items.Add(drExcel[drExcel.GetName(0)].ToString().Replace('_', ' '));


                        for (var h = 1; h < drExcel.FieldCount; h++)
                            if (
                                drExcel.GetName(h).ToString() == "image" ||
                                drExcel.GetName(h).ToString() == "photo" ||
                                drExcel.GetName(h).ToString() == "Logo" ||
                                drExcel.GetName(h).ToString() == "ID")
                            {
                            }
                            else
                            {
                                var checker = "";
                                if (drExcel.GetName(h).ToString() == "Miles Per Gallon" ||
                                    drExcel.GetName(h).ToString() == "Litres Per 100 km" ||
                                    drExcel.GetName(h).ToString() == "Fine Amount" ||
                                    drExcel.GetName(h).ToString() == "Vatable Total" ||
                                    drExcel.GetName(h).ToString() == "Subtotal" ||
                                    drExcel.GetName(h).ToString() == "Tax" ||
                                    drExcel.GetName(h).ToString() == "Total" ||
                                    drExcel.GetName(h).ToString() == "Debit Amount" ||
                                    drExcel.GetName(h).ToString() == "Total VAT" ||
                                    drExcel.GetName(h).ToString() == "Vatable Total" ||
                                    drExcel.GetName(h).ToString() == "Paid Amount" ||
                                    drExcel.GetName(h).ToString() == "Balance" ||
                                    drExcel.GetName(h).ToString() == "Consumed Qty" ||
                                    drExcel.GetName(h).ToString() == "Initial Qty" ||
                                    drExcel.GetName(h).ToString() == "Qty" ||
                                    drExcel.GetName(h).ToString() == "VAT AMOUNT" ||
                                    drExcel.GetName(h).ToString() == "Debit Amount" ||
                                    drExcel.GetName(h).ToString() == "Netchange Debit" ||
                                    drExcel.GetName(h).ToString() == "Netchange Credit" ||
                                    drExcel.GetName(h).ToString() == "Balance Debit" ||
                                    drExcel.GetName(h).ToString() == "Balance Credit" ||
                                    drExcel.GetName(h).ToString() == "Credit Amount" ||
                                    drExcel.GetName(h).ToString() == "VAT%" ||
                                    drExcel.GetName(h).ToString() == "Total Cost" ||
                                    drExcel.GetName(h).ToString() == "VAT Amount" ||
                                    drExcel.GetName(h).ToString() == "Unit Cost" ||
                                    drExcel.GetName(h).ToString() == "Balance Credit" ||
                                    drExcel.GetName(h).ToString() == "Balance Debit" ||
                                    drExcel.GetName(h).ToString() == "Net Change Credit" ||
                                    drExcel.GetName(h).ToString() == "Net Change Debit" ||
                                    drExcel.GetName(h).ToString() == "Net Amount" ||
                                    drExcel.GetName(h).ToString() == "Receipt Amount" ||
                                    drExcel.GetName(h).ToString() == "Total Deductions" ||
                                    drExcel.GetName(h).ToString() == "Pension" ||
                                    drExcel.GetName(h).ToString() == "NHIF" ||
                                    drExcel.GetName(h).ToString() == "NSSF" ||
                                    drExcel.GetName(h).ToString() == "PAYE" || drExcel.GetName(h).ToString() == "MPR" ||
                                    drExcel.GetName(h).ToString() == "Tax Charged" ||
                                    drExcel.GetName(h).ToString() == "Net Pay" ||
                                    drExcel.GetName(h).ToString() == "Taxable Pay" ||
                                    drExcel.GetName(h).ToString() == "Gross Pay" ||
                                    drExcel.GetName(h).ToString() == "Basic Pay" ||
                                    drExcel.GetName(h).ToString() == "Amount" ||
                                    drExcel.GetName(h).ToString() == "Cost" ||
                                    drExcel.GetName(h).ToString() == "Price" ||
                                    drExcel.GetName(h).ToString() == "Day Weighting")
                                {
                                    if (drExcel[drExcel.GetName(h)].ToString() == "")
                                        checker = "0";
                                    else
                                        checker = drExcel[drExcel.GetName(h)].ToString();

                                    lv.SubItems.Add(string.Format("{0:0.00}", double.Parse(checker)));
                                }
                                else
                                {
                                    lv.SubItems.Add(drExcel[drExcel.GetName(h)].ToString());
                                }
                            }
                    }
                }

                for (var i = 1; i < lsvMain.Columns.Count; i++)
                    lsvMain.Columns[i].Width = -2;
                cnExcel.Close();
                // Cursor.Current = Cursors.Arrow;
            }
            catch (Exception)
            {
                //  classPublicclass.showMessage(ex.Message);
            }

          
        }

        public static bool GrantAccess(string fullPath) // make path available for writing
        {
            try
            {
                var dInfo = new DirectoryInfo(fullPath);
                var dSecurity = dInfo.GetAccessControl();
                dSecurity.AddAccessRule(new FileSystemAccessRule("everyone", FileSystemRights.FullControl,
                    InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit,
                    PropagationFlags.NoPropagateInherit, AccessControlType.Allow));
                dInfo.SetAccessControl(dSecurity);
            }
            catch (Exception)
            {
            }

            // return true;

            try
            {
                var dInfo = new DirectoryInfo(fullPath);
                var dSecurity = dInfo.GetAccessControl();
                dSecurity.AddAccessRule(new FileSystemAccessRule(
                    new SecurityIdentifier(WellKnownSidType.WorldSid, null), FileSystemRights.FullControl,
                    InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit,
                    PropagationFlags.NoPropagateInherit, AccessControlType.Allow));
                dInfo.SetAccessControl(dSecurity);
            }
            catch (Exception)
            {
            }

            return true;
        }

        public static void Minimizeform(Form myform)
           {
               myform.WindowState = FormWindowState.Minimized;
               myform.FormBorderStyle = FormBorderStyle.FixedSingle;
           }

        public static void Formresizingmdiparent(Form myform, ToolStripButton toolStripButton2)
           {
               var img = new FrmSystemImages();

               if (myform.WindowState == FormWindowState.Normal)
               {
                   // myform.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                   myform.FormBorderStyle = FormBorderStyle.None;
                   toolStripButton2.Image = img.maximum.Image;
                   toolStripButton2.ToolTipText = "Maximize";
               }

               if (myform.WindowState == FormWindowState.Maximized)
               {
                   //myform.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
                   myform.FormBorderStyle = FormBorderStyle.None;
                   toolStripButton2.Image = img.normal.Image;
                   toolStripButton2.ToolTipText = "Restore";
               }
           }

        public static void Resizeformmdiparent(Form myform, ToolStripButton toolStripButton2)
           {
               var img = new FrmSystemImages();
               if (myform.WindowState == FormWindowState.Maximized)
               {
                   myform.WindowState = FormWindowState.Normal;
                   toolStripButton2.Image = img.maximum.Image;
                   toolStripButton2.ToolTipText = "Maximize";
                   // myform.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
               }
               else
               {
                   myform.WindowState = FormWindowState.Maximized;
                   toolStripButton2.Image = img.normal.Image;
                   toolStripButton2.ToolTipText = "Restore";
                   //  myform.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
                   myform.FormBorderStyle = FormBorderStyle.None;
               }
           }

        public static void Fillcombowithcolumnames(ComboBox comboBox1, string Tablename, string condition)
        {
            try
            {
                comboBox1.Items.Clear();

                var sbConnCOMBOx = new StringBuilder();
                sbConnCOMBOx.Append(ClassDatabaseConnection.cnn1);
                sbConnCOMBOx.Append(";Extended Properties=");
                sbConnCOMBOx.Append(Convert.ToChar(34));
                sbConnCOMBOx.Append(Convert.ToChar(34));
                var cnExcelCOMBOx = new OleDbConnection(sbConnCOMBOx.ToString());
                cnExcelCOMBOx.Open();
                var sbSQLCOMBOx = new StringBuilder();
                sbSQLCOMBOx.Append(string.Format("Select * FROM [{0}]", Tablename));
                var cmdExcelCOMBOx = new OleDbCommand(sbSQLCOMBOx.ToString(), cnExcelCOMBOx);
                var drExcelCOMBOx = cmdExcelCOMBOx.ExecuteReader();
                for (var i = 0; i < drExcelCOMBOx.FieldCount; i++) 

                    // filter out unwanted columns for populating combobox with column names
                    if (drExcelCOMBOx.GetName(i).ToString() == "ID" ||
                        drExcelCOMBOx.GetName(i).ToString() == "Photo" ||
                        drExcelCOMBOx.GetName(i).ToString() == "Image" ||
                        drExcelCOMBOx.GetName(i).ToString() == "photo" ||          
                        drExcelCOMBOx.GetName(i).ToString() == "password" ||      
                        drExcelCOMBOx.GetName(i).ToString() == "Password"
                    )
                    {
                    }
                    else
                    {
                       comboBox1.Items.Add(drExcelCOMBOx.GetName(i).ToString());
                    }

                cnExcelCOMBOx.Close();
            }
            catch (Exception)
            {
            }
        }

        public static void Generatereport(string title, string extension, string sql, string widthn)
        {
            var path = ClassPublicclass.GetTemporaryDirectory("DE_Reports");            

               
            var rnd = new Random();
                var surfix = rnd.Next(52); // randomly generate a number and append to title to void error if title (filename) already exixts

                
                var docname = title + surfix.ToString();
                
                 var Reports = "DE_Reports";
                    GrantAccess(@path + "\\DE_Reports\\" + docname + "." + extension);
                    File.Delete(@path + "\\DE_Reports\\" + docname + "." + extension);
                    TextWriter tw = new StreamWriter(string.Format("{0}\\{1}\\{2}.{3}", path, Reports, docname, extension));
                    tw.WriteLine(
                        "<style type=\"text/css\" media=\"print\">  @page land { size: landscape; }</style><htm><header></header><body class=\"land\">");
                    //tw.WriteLine("<p><img src=@"+ path +"\\INTG_Images\\clogo.bmp"+"/></p>");

                    // close the stream
                    //tw.Close();
                    var companyNameLbl = new Label();
                    var comptellbl = new Label();
                    var mobilelbl = new Label();
                    var compaddlbl = new Label();
                    var companylogo1lbl = new Label();
                    var websitelbl = new Label();
                    var emaillbl = new Label();
                    var faxlbl = new Label();
                    
                    companyNameLbl.Text = "MY COMPANY";
               


                    // Cursor.Current = Cursors.WaitCursor;
                    var sbConnCOMBOx = new StringBuilder();
                    sbConnCOMBOx.Append(ClassDatabaseConnection.cnn1);
                    sbConnCOMBOx.Append(";Extended Properties=");
                    sbConnCOMBOx.Append(Convert.ToChar(34));
                    sbConnCOMBOx.Append(Convert.ToChar(34));
                    var cnExcelCOMBOx = new OleDbConnection(sbConnCOMBOx.ToString());
                    cnExcelCOMBOx.Open();
                    var sbSQLCOMBOx = new StringBuilder();
                    sbSQLCOMBOx.Append(sql);
                    var cmdExcelCOMBOx = new OleDbCommand(sbSQLCOMBOx.ToString(), cnExcelCOMBOx);
                    var drExcelCOMBOx = cmdExcelCOMBOx.ExecuteReader();

                    var contentTable = new PdfPTable(drExcelCOMBOx.FieldCount);
                    var mycolspan = 3;
                    if (drExcelCOMBOx.FieldCount > 4) mycolspan = drExcelCOMBOx.FieldCount - 2;

                    // header
                    tw.WriteLine("<table><tr><td colspan=\"" + mycolspan + "\">" + companyNameLbl.Text +
                                 " </td><td colspan=\"2\">" + DateTime.Now.Date.ToString(ClassMainclass.Dateformat) +
                                 "  " + DateTime.Now.ToShortTimeString() + " </td></tr>");
                    tw.WriteLine("<p>&nbsp;</p>");
                    tw.WriteLine("<tr><td> Address " + compaddlbl.Text + " </td>");
                    tw.WriteLine("<td> Telephone " + comptellbl.Text + " </td>");
                    tw.WriteLine("<td> Website " + websitelbl.Text + " </td></tr>");
                    tw.WriteLine("<tr><td> Mobile " + mobilelbl.Text + " </td>");
                    tw.WriteLine("<td> Email " + emaillbl.Text + " </td>");
                    tw.WriteLine("<td> Fax " + faxlbl.Text + " </td></tr></table>");
                    tw.WriteLine("<p>&nbsp;</p>");
                    tw.WriteLine("<p><b>" + title + "</b></p>");
                    tw.WriteLine("<table border=\"1\" bordercolor=\"#003366\"  cellspacing=\"0\" width=\"100%\">");
                    tw.WriteLine("<tr>");
                    for (var i = 0; i < drExcelCOMBOx.FieldCount; i++)
                        if (drExcelCOMBOx.GetName(i).ToUpper() == "ID")
                            tw.WriteLine("<td bgcolor=\"#003366\"><font color=\"#FFFFFF\"><b>No</b></font></td>");
                        else
                            tw.WriteLine("<td bgcolor=\"#003366\"><font color=\"#FFFFFF\"><b>" +
                                         drExcelCOMBOx.GetName(i).ToString().Replace('_', ' ') + " </b></font></td>");

                    tw.Write("</tr>");
                    cnExcelCOMBOx.Close();
                    var sbConnCOMBO = new StringBuilder();
                    sbConnCOMBO.Append(ClassDatabaseConnection.cnn1);
                    sbConnCOMBO.Append(";Extended Properties=");
                    sbConnCOMBO.Append(Convert.ToChar(34));
                    sbConnCOMBO.Append(Convert.ToChar(34));

                    var cnExcelCOMBO = new OleDbConnection(sbConnCOMBO.ToString());
                    cnExcelCOMBO.Open();
                    var sbSQLCOMBO = new StringBuilder();
                    sbSQLCOMBO.Append(sql);
                    var cmdExcelCOMBO = new OleDbCommand(sbSQLCOMBO.ToString(), cnExcelCOMBO);
                    var drExcelCOMBO = cmdExcelCOMBO.ExecuteReader();

                    while (drExcelCOMBO.Read())
                    {
                        tw.Write("<tr>");

                        for (var K = 0; K < drExcelCOMBO.FieldCount; K++) tw.Write("<td>" + drExcelCOMBO[drExcelCOMBO.GetName(K)].ToString() + "</d>");

                        tw.Write("</tr>");
                    }

                    cnExcelCOMBO.Close();


                    tw.WriteLine("</table>");
                    tw.WriteLine("</body></htm>");

                    tw.Close();

                    try
                    {
                        Process.Start(@path + "\\DE_Reports\\" + docname + "." + extension);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("The document may be open, please close it and try again"+ex.ToString());
                    }
                
                

                // Microsoft.Office.Interop.Word.Document newDoc = new Microsoft.Office.Interop.Word.Document();
            }
        }    
}
