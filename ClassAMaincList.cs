using System;
using System.Data.OleDb;
using System.Text;
using System.Windows.Forms;

namespace DataExplorer2
{
    internal class ClassAMaincList
    {
        

        public static string Mycolumns(string tablename)
        {
            var mycolumns = "";

            return mycolumns;
        }

        public static string Columncount(string tablename)
        {          
            var cocount = new ListView();
            ClassPublicclass.Filllistnoid(tablename);  //populate listview and ignore the ID column
            var colcount = cocount.Items.Count.ToString();
            return colcount;
        }

        public static string Listsql(string tablename, string ishort, string toptxt)
        {
            ClassMainclass.listselect = toptxt;
            var    mysql = " select   "+ClassMainclass.listselect+"  * from " + tablename;
            return mysql;
        }

        public static string FilterList(string tablename, TextBox textBox1, string formText, string toptxt)
        {
            switch (toptxt)
            {
                case "ALL": ClassMainclass.listselect = ""; break;
                case "": ClassMainclass.listselect = ""; break;
                default: ClassMainclass.listselect = "TOP " + toptxt + " "; break;
            }
            
            var mysql = "";
            if (textBox1.Text == "")
            {
                mysql = "";
            }
            else
            {
                       mysql = " ";


                            var mysql1 = "(";
                            try
                            {
                                var sbConnCombOx = new StringBuilder();
                                sbConnCombOx.Append(ClassDatabaseConnection.cnn1);
                                sbConnCombOx.Append(";Extended Properties=");
                                sbConnCombOx.Append(Convert.ToChar(34));
                                sbConnCombOx.Append(Convert.ToChar(34));
                                var cnExcelCombOx = new OleDbConnection(sbConnCombOx.ToString());
                                cnExcelCombOx.Open();
                                var sbSqlcombOx = new StringBuilder();
                                sbSqlcombOx.Append(string.Format("Select * FROM [{0}]", tablename));
                                var cmdExcelCombOx = new OleDbCommand(sbSqlcombOx.ToString(), cnExcelCombOx);
                                var drExcelCombOx = cmdExcelCombOx.ExecuteReader();
                                for (var i = 0; i < drExcelCombOx.FieldCount; i++)

                                if (
                                drExcelCombOx.GetName(i).ToString().ToUpper() == "IMAGE" ||
                                drExcelCombOx.GetName(i).ToString().ToUpper() == "PHOTO" ||  
                                drExcelCombOx.GetName(i).ToString().ToUpper() == "PASSWORD" ||
                                drExcelCombOx.GetName(i).ToString().ToUpper() == "PIC" ||
                                drExcelCombOx.GetName(i).ToString().ToUpper() == "PICTURE" ||
                                drExcelCombOx.GetName(i).ToString().ToUpper() == "PASSWORD" ||
                                drExcelCombOx.GetName(i).ToString().ToUpper() == "LOGO")
                                 {
                                 }
                                    else
                                    {

                                        mysql1 = mysql1 + "  [" + drExcelCombOx.GetName(i) + "] LIKE '%" +
                                                 textBox1.Text + "%' OR ";
                                    }
             
                                cnExcelCombOx.Close();
                            }
                            catch (Exception)
                            {
                            }

                            var rectify = mysql1 + "--";

                            mysql1 = rectify.Replace("OR --", ")");

                            mysql = " WHERE   " + mysql1;

                
            }
           // MessageBox.Show(mysql);
            return mysql;

          
        }
    }
}