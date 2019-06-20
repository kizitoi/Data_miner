using System;
using System.ComponentModel;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Data.SqlClient;

namespace DataExplorer2
{
    public partial class FrmSystemMainList : Form
    {
        private ClassListSorter lvwColumnSorter;

        public FrmSystemMainList()
        {
            InitializeComponent();
            // Create an instance of a ListView column sorter and assign it 
            // to the ListView control.
            lvwColumnSorter = new ClassListSorter();
            listView2.ListViewItemSorter = lvwColumnSorter;
        }

        public string fullsql = " ";
        public string shortsql = " ";
        public TextBox txtID = new TextBox();
        public bool ismeter = false;
            
   

        private void Colorform()
        {
            toolStrip4.BackColor = System.Drawing.Color.Blue;
            toolStrip1.BackColor = System.Drawing.Color.Blue;
            BackColor = System.Drawing.Color.Gray;
        }


     
        private void FrmSystemMainList_Load(object sender, EventArgs e)
        {
            //Form Load
            var img = new FrmSystemImages();
            txtsimpleSearch.KeyPress += new KeyPressEventHandler(TbPassword_KeyPress);
            FormBorderStyle = FormBorderStyle.None;
           // WindowState = FormWindowState.Maximized;

            this.comboBox2.KeyPress += ClassPublicclass.Numbersonly;
         

            // Ensure that the view is set to show details.
            listView2.View = View.Details;

     
            refresh.Image = img.refrech.Image;
            pdf.Image = img.pdf.Image;
            excel.Image = img.excel.Image;
            word.Image = img.word.Image;
         
          
          
            btnMinimize.Image = img.minimize.Image;
            btnMaximize.Image = img.maximum.Image;
            btnClose.Image = img.close.Image;
          
        

            btnMaximize.Image = img.normal.Image;

            BackColor = ClassMainclass.FormColor;
            toolStrip4.BackColor = ClassMainclass.ToolstripColor;
            toolStrip1.BackColor = ClassMainclass.ToolstripColor;
            button1.BackColor = ClassMainclass.ToolstripColor;
            panel4.BackColor = ClassMainclass.ToolstripColor;
            panel5.BackColor = ClassMainclass.ToolstripColor;

            panel14.BackColor = ClassMainclass.ToolstripColor;
            button2.BackColor = ClassMainclass.ToolstripColor;
            button3.BackColor = ClassMainclass.ToolstripColor;
            toolStripLabel1.ForeColor = ClassMainclass.ToolstripFontColor;


            panel11.BackColor = ClassMainclass.ToolstripColor;
            label4.BackColor = ClassMainclass.ToolstripColor;
            label4.ForeColor = ClassMainclass.ToolstripFontColor;

            button4.BackColor = ClassMainclass.ToolstripColor;
            button4.ForeColor = ClassMainclass.ToolstripFontColor;

            button5.BackColor = ClassMainclass.ToolstripColor;
            button5.ForeColor = ClassMainclass.ToolstripFontColor;

            button6.BackColor = ClassMainclass.ToolstripColor;
            button6.ForeColor = ClassMainclass.ToolstripFontColor;

            FormBorderStyle = FormBorderStyle.None;     
            toolStripLabel1.Text = Text + " ";


            if (comboBox2.Text == "ALL")
            {
                ClassMainclass.listselect = " ";
            }
            else
            {
                ClassMainclass.listselect = "TOP " + comboBox2.Text + " ";
            }

            try
            {
                fullsql = ClassAMaincList.Listsql(txtTableName.Text, " ID,", ClassMainclass.listselect);
            }
            catch { }

   
            try
            {
                shortsql = ClassAMaincList.Listsql(txtTableName.Text, " ", ClassMainclass.listselect);
            }
            catch { }
            

            ClassPublicclass.Selectfromtablecountlbl(label8, txtTableName.Text, "ID", "");
            txtTableName.Text = cmbTableName.Text;
            txtServerName.Text = listboxSQLServerInstances.Text;
            txtDatabaseName.Text = listboxSQLServerDatabaseInstances.Text;
        }

        public Label customerNo = new Label();

        public void Filllist()
        {
            ClassPublicclass.Filllist(listView2, fullsql);
        
        }

        private void ToolStripButton27_Click(object sender, EventArgs e)
        {
            Refreshlist();
        }

        public void Refreshlist()
        {
                      
            ClassPublicclass.Filllist(listView2, fullsql);
            txtsimpleSearch.Text = "";
           
        }

        public void Refreshme()
        {
            ClassPublicclass.Filllist(listView2, fullsql);
            txtsimpleSearch.Text = "";
            label2.Text = listView2.Items.Count.ToString();
           
        }

     
        private string editview = "";
        

        private void ToolStripButton26_Click(object sender, EventArgs e)
        {
            var orientation = "L";  // LANDSCAPE

            if (radioButton1.Checked == false)
            { orientation = "P"; //PORTRAIT
            
            }

            if (panel6.Visible == true)
                ClassPrintPDF.Generatereportpdf1(Text, "pdf", filteredprint, orientation);
            else
                ClassPrintPDF.Generatereportpdf1(Text, "pdf", shortsql, orientation);
        }

        private void ToolStripButton25_Click(object sender, EventArgs e)
        {
            if (panel6.Visible == true)
                ClassPublicclass.Generatereport(Text, "xls", filteredprint, "100%");
            else
                ClassPublicclass.Generatereport(Text, "xls", shortsql, "100%");
        }

     

        private void TbPassword_KeyPress(object sender, KeyPressEventArgs e)
        {
            char keyChar;
            keyChar = e.KeyChar;

            if (keyChar == 13)
            {
                Search_Go();
                e.Handled = true;
            }

           
        }

        private void Search_Go()
        {


            if (comboBox2.Text == "ALL")
            {
                ClassMainclass.listselect = " ";
            }
            else
            {
                ClassMainclass.listselect = "TOP " + comboBox2.Text + " ";
            }

            fullsql = ClassAMaincList.Listsql(txtTableName.Text, " ID,", ClassMainclass.listselect);


            if (txtsimpleSearch.Text == "")
            {
                editview = ClassAMaincList.FilterList(txtTableName.Text, txtsimpleSearch, Text, ClassMainclass.listselect);
                ClassPublicclass.Filllist(listView2, fullsql + editview);
                shortsql = fullsql + editview;
            }
            else
            {
                editview = ClassAMaincList.FilterList(txtTableName.Text, txtsimpleSearch, Text, ClassMainclass.listselect);
                ClassPublicclass.Filllist(listView2, fullsql + editview);
                shortsql = fullsql + editview;
            }


            label2.Text = listView2.Items.Count.ToString();
            


        }

        private void Button1_Click(object sender, EventArgs e)
        {
            //filtermenow();

            Search_Go();


        }


        private void Filtermenow()
        {
            try
            {
                switch (comboBox2.Text)
                {
                    case "ALL": ClassMainclass.listselect = ""; break;
                    case "": ClassMainclass.listselect = ""; break;

                    default: ClassMainclass.listselect = "TOP " + comboBox2.Text + " "; break;
                }
                         
               
                fullsql = ClassAMaincList.Listsql(txtTableName.Text, " ID,", ClassMainclass.listselect);
                editview = ClassAMaincList.FilterList(txtTableName.Text, txtsimpleSearch, Text, ClassMainclass.listselect);

                fullsql = fullsql + editview;
                ClassPublicclass.Filllist(listView2, fullsql + editview);
                shortsql = fullsql + editview;
                label2.Text = listView2.Items.Count.ToString();

               

            }
            catch { }
        }

     

        private void Button2_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text == "ALL")
            {
                ClassMainclass.listselect = "";
            }
            else
            {
                ClassMainclass.listselect = "TOP " + comboBox2.Text + " ";
            }

            txtsimpleSearch.Text = "";
           
            editview = ClassAMaincList.FilterList(txtTableName.Text, txtsimpleSearch, Text, ClassMainclass.listselect);
            ClassPublicclass.Filllist(listView2, fullsql + editview);
            shortsql = fullsql + editview;
           
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ToolStripButton1_Click(object sender, EventArgs e)

        {
            Close();
        }

        private void ToolStripButton2_Click(object sender, EventArgs e)
        {
            ClassPublicclass.Resizeformmdiparent(this, btnMaximize);
            FormBorderStyle = FormBorderStyle.None;
        }

        private void ToolStripButton3_Click(object sender, EventArgs e)
        {
            ClassPublicclass.Minimizeform(this);
        }

        private void FrmSystemMainList_SizeChanged(object sender, EventArgs e)
        {
        }

        //const and dll functions for moving form
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;

        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd,
            int Msg, int wParam, int lParam);

        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();

        private void ToolStrip4_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }


        private void ToolStrip4_DoubleClick(object sender, EventArgs e)
        {
        }

        private void FrmSystemMainList_ResizeEnd(object sender, EventArgs e)
        {
            ClassPublicclass.Formresizingmdiparent(this, btnMaximize);
            FormBorderStyle = FormBorderStyle.None;
        }

        private void Word_Click(object sender, EventArgs e)
        {
            if (panel6.Visible == true)
                ClassPublicclass.Generatereport(Text, "doc", filteredprint, "100%");
            else
                ClassPublicclass.Generatereport(Text, "doc", shortsql, "100%");
        }

        private void ListView2_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order ==System.Windows.Forms.SortOrder.Ascending)
                    lvwColumnSorter.Order = System.Windows.Forms.SortOrder.Descending;
                else
                    lvwColumnSorter.Order = System.Windows.Forms.SortOrder.Ascending;
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = System.Windows.Forms.SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            listView2.Sort();
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
            if (comboBox2.Text == "ALL")
            {
                ClassMainclass.listselect = " ";
            }
            else
            {
                ClassMainclass.listselect = "TOP " + comboBox2.Text + " ";
            }
         
            fullsql = ClassAMaincList.Listsql(txtTableName.Text, " ID,", ClassMainclass.listselect);          


            if (txtsimpleSearch.Text == "")
            {
                editview = ClassAMaincList.FilterList(txtTableName.Text, txtsimpleSearch, Text, ClassMainclass.listselect);
                ClassPublicclass.Filllist(listView2, fullsql + editview);
                shortsql = fullsql + editview;               
            }
                else
                {
                    editview = ClassAMaincList.FilterList(txtTableName.Text, txtsimpleSearch, Text, ClassMainclass.listselect);
                    ClassPublicclass.Filllist(listView2, fullsql + editview);
                    shortsql = fullsql + editview;
                }

              
                label2.Text = listView2.Items.Count.ToString();
            }
        }

       

        private void ListView2_Validating(object sender, CancelEventArgs e)
        {
            label2.Text = listView2.Items.Count.ToString();
        }

        private void ListView2_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            e.DrawDefault = true;
        }

  
        private void LinkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
           
            if (panel6.Visible == false)
            {
                panel6.Visible = true;
                ClassPublicclass.Fillcombowithcolumnames(comboBox1, txtTableName.Text, "");
                if (panel9.Controls.Count == 2)
                    Adjustme();
            }
            else
            {
                panel6.Visible = false;
            }
        }

        private void LinkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private int _i = 0;

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
            }
        }

        private string txt0 = "",
            txt1 = "",
            txt2 = "",
            txt3 = "",
            txt4 = "",
            txt5 = "",
            txt6 = "",
            txt7 = "",
            txt8 = "",
            txt9 = "",
            txt10 = "",
            txt11 = "",
            txt12 = "",
            txt13 = "",
            txt14 = "",
            txt15 = "",
            txt16 = "";

        private string cmb0 = "",
            cmb1 = "",
            cmb2 = "",
            cmb3 = "",
            cmb4 = "",
            cmb5 = "",
            cmb6 = "",
            cmb7 = "",
            cmb8 = "",
            cmb9 = "",
            cmb10 = "",
            cmb11 = "",
            cmb12 = "",
            cmb13 = "",
            cmb14 = "",
            cmb15 = "",
            cmb16 = "";

        private string nor0 = "",
            nor1 = "",
            nor2 = "",
            nor3 = "",
            nor4 = "",
            nor5 = "",
            nor6 = "",
            nor7 = "",
            nor8 = "",
            nor9 = "",
            nor10 = "",
            nor11 = "",
            nor12 = "",
            nor13 = "",
            nor14 = "",
            nor15 = "",
            nor16 = "";

        private string lln0 = "",
            lln1 = "",
            lln2 = "",
            lln3 = "",
            lln4 = "",
            lln5 = "",
            lln6 = "",
            lln7 = "",
            lln8 = "",
            lln9 = "",
            lln10 = "",
            lln11 = "",
            lln12 = "",
            lln13 = "",
            lln14 = "",
            lln15 = "",
            lln16 = "";


        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            switch (((ComboBox) sender).Name)
            {
                case "cmb0":
                    cmb0 = ((ComboBox) sender).Text;
                    break;
                case "cmb1":
                    cmb1 = ((ComboBox) sender).Text;
                    break;
                case "cmb2":
                    cmb2 = ((ComboBox) sender).Text;
                    break;
                case "cmb3":
                    cmb3 = ((ComboBox) sender).Text;
                    break;
                case "cmb4":
                    cmb4 = ((ComboBox) sender).Text;
                    break;
                case "cmb5":
                    cmb5 = ((ComboBox) sender).Text;
                    break;
                case "cmb6":
                    cmb6 = ((ComboBox) sender).Text;
                    break;
                case "cmb7":
                    cmb7 = ((ComboBox) sender).Text;
                    break;
                case "cmb8":
                    cmb8 = ((ComboBox) sender).Text;
                    break;
                case "cmb9":
                    cmb9 = ((ComboBox) sender).Text;
                    break;
                case "cmb10":
                    cmb10 = ((ComboBox) sender).Text;
                    break;
                case "cmb11":
                    cmb11 = ((ComboBox) sender).Text;
                    break;
                case "cmb12":
                    cmb12 = ((ComboBox) sender).Text;
                    break;
                case "cmb13":
                    cmb13 = ((ComboBox) sender).Text;
                    break;
                case "cmb14":
                    cmb14 = ((ComboBox) sender).Text;
                    break;
                case "cmb15":
                    cmb15 = ((ComboBox) sender).Text;
                    break;
                case "cmb16":
                    cmb16 = ((ComboBox) sender).Text;
                    break;

                case "lln0":
                    lln0 = ((ComboBox) sender).Text;
                    break;
                case "lln1":
                    lln1 = ((ComboBox) sender).Text;
                    break;
                case "lln2":
                    lln2 = ((ComboBox) sender).Text;
                    break;
                case "lln3":
                    lln3 = ((ComboBox) sender).Text;
                    break;
                case "lln4":
                    lln4 = ((ComboBox) sender).Text;
                    break;
                case "lln5":
                    lln5 = ((ComboBox) sender).Text;
                    break;
                case "lln6":
                    lln6 = ((ComboBox) sender).Text;
                    break;
                case "lln7":
                    lln7 = ((ComboBox) sender).Text;
                    break;
                case "lln8":
                    lln8 = ((ComboBox) sender).Text;
                    break;
                case "lln9":
                    lln9 = ((ComboBox) sender).Text;
                    break;
                case "lln10":
                    lln10 = ((ComboBox) sender).Text;
                    break;
                case "lln11":
                    lln11 = ((ComboBox) sender).Text;
                    break;
                case "lln12":
                    lln12 = ((ComboBox) sender).Text;
                    break;
                case "lln13":
                    lln13 = ((ComboBox) sender).Text;
                    break;
                case "lln14":
                    lln14 = ((ComboBox) sender).Text;
                    break;
                case "lln15":
                    lln15 = ((ComboBox) sender).Text;
                    break;
                case "lln16":
                    lln16 = ((ComboBox) sender).Text;
                    break;

                case "nor0":
                    nor0 = ((ComboBox) sender).Text;
                    break;
                case "nor1":
                    nor1 = ((ComboBox) sender).Text;
                    break;
                case "nor2":
                    nor2 = ((ComboBox) sender).Text;
                    break;
                case "nor3":
                    nor3 = ((ComboBox) sender).Text;
                    break;
                case "nor4":
                    nor4 = ((ComboBox) sender).Text;
                    break;
                case "nor5":
                    nor5 = ((ComboBox) sender).Text;
                    break;
                case "nor6":
                    nor6 = ((ComboBox) sender).Text;
                    break;
                case "nor7":
                    nor7 = ((ComboBox) sender).Text;
                    break;
                case "nor8":
                    nor8 = ((ComboBox) sender).Text;
                    break;
                case "nor9":
                    nor9 = ((ComboBox) sender).Text;
                    break;
                case "nor10":
                    nor10 = ((ComboBox) sender).Text;
                    break;
                case "nor11":
                    nor11 = ((ComboBox) sender).Text;
                    break;
                case "nor12":
                    nor12 = ((ComboBox) sender).Text;
                    break;
                case "nor13":
                    nor13 = ((ComboBox) sender).Text;
                    break;
                case "nor14":
                    nor14 = ((ComboBox) sender).Text;
                    break;
                case "nor15":
                    nor15 = ((ComboBox) sender).Text;
                    break;
                case "nor16":
                    nor16 = ((ComboBox) sender).Text;
                    break;
            }
        }

        private void TextBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // textBox3.Text = textBox3.Text + "  " + ((TextBox)sender).Text;
            // ((TextBox)sender).Enabled = false;

            switch (((TextBox) sender).Name)
            {
                case "txt0":
                    txt0 = ((TextBox) sender).Text;
                    break;
                case "txt1":
                    txt1 = ((TextBox) sender).Text;
                    break;
                case "txt2":
                    txt2 = ((TextBox) sender).Text;
                    break;
                case "txt3":
                    txt3 = ((TextBox) sender).Text;
                    break;
                case "txt4":
                    txt4 = ((TextBox) sender).Text;
                    break;
                case "txt5":
                    txt5 = ((TextBox) sender).Text;
                    break;
                case "txt6":
                    txt6 = ((TextBox) sender).Text;
                    break;
                case "txt7":
                    txt7 = ((TextBox) sender).Text;
                    break;
                case "txt8":
                    txt8 = ((TextBox) sender).Text;
                    break;
                case "txt9":
                    txt9 = ((TextBox) sender).Text;
                    break;
                case "txt10":
                    txt10 = ((TextBox) sender).Text;
                    break;
                case "txt11":
                    txt11 = ((TextBox) sender).Text;
                    break;
                case "txt12":
                    txt12 = ((TextBox) sender).Text;
                    break;
                case "txt13":
                    txt13 = ((TextBox) sender).Text;
                    break;
                case "txt14":
                    txt14 = ((TextBox) sender).Text;
                    break;
                case "txt15":
                    txt15 = ((TextBox) sender).Text;
                    break;
                case "txt16":
                    txt16 = ((TextBox) sender).Text;
                    break;
            }
        }

        private void Adjustme()
        {
           
            var a = new Panel();
            panel9.Controls.Add(a);
            a.Dock = DockStyle.Top;
            a.BringToFront();
            a.Height = panel8.Height;
            a.Name = "pan" + _i.ToString();
            // a.BackColor = Color.Blue;          
            var combo = new ComboBox
            {
                Width = comboBox1.Width,
                Name = "cmb" + _i.ToString()
            };
            ClassPublicclass.Fillcombowithcolumnames(combo, txtTableName.Text, "");
            //combo.Location = new Point(button1.Location.X, button1.Location.Y + combo.Height * _i);
            combo.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
            combo.Dock = DockStyle.Left;


            var nor = new ComboBox
            {
                Name = "nor" + _i.ToString()
            };
            nor.Items.Add("OR");
            nor.Items.Add("AND");
            nor.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
            nor.Text = "=";
            nor.SelectedIndex = 0;
            nor.Dock = DockStyle.Left;
            nor.Width = label5.Width;
            a.Controls.Add(nor);

            var ntxt = new TextBox
            {
                Name = "txt" + _i.ToString(),
                Dock = DockStyle.Left
            };
            ntxt.TextChanged += new EventHandler(TextBox_SelectedIndexChanged);
            ntxt.Width = textBox2.Width;
            a.Controls.Add(ntxt);

            var lln = new ComboBox
            {
                Name = "lln" + _i.ToString()
            };
            lln.Items.Add("=");
            lln.Items.Add("LIKE");
            lln.Items.Add("<");
            lln.Items.Add(">");
            lln.Items.Add("<=");
            lln.Items.Add(">=");
            lln.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
            lln.Text = "=";
            lln.SelectedIndex = 0;
            lln.Dock = DockStyle.Left;
            lln.Width = label5.Width;
            a.Controls.Add(lln);
            a.Controls.Add(combo);        
            _i++;
           
        }

        private void Button5_Click(object sender, EventArgs e)
        {
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            var mysql = "";
            filteredprint = "";

            // string sign = "LIKE";
            var likesign = "%";


            if (cmb0 != "")
            {
                if (lln0 != "LIKE") likesign = "";

                mysql = "[" + cmb0 + "] " + lln0 + " '" + likesign + txt0 + likesign + "' ";
            }

            if (cmb1 != "")
            {
                if (lln1 != "LIKE") likesign = "";

                mysql = mysql + " " + nor0 + " [" + cmb1 + "] " + lln1 + "'" + likesign + txt1 + likesign + "'";
            }

            if (cmb2 != "")
            {
                if (lln2 != "LIKE") likesign = "";

                mysql = mysql + " " + nor1 + " [" + cmb2 + "] " + lln2 + "'" + likesign + txt2 + likesign + "'";
            }

            if (cmb3 != "")
            {
                if (lln3 != "LIKE") likesign = "";

                mysql = mysql + " " + nor2 + " [" + cmb3 + "] " + lln3 + "'" + likesign + txt3 + likesign + "'";
            }

            if (cmb4 != "")
            {
                if (lln4 != "LIKE") likesign = "";

                mysql = mysql + " " + nor3 + " [" + cmb4 + "] " + lln4 + "'" + likesign + txt4 + likesign + "'";
            }

            if (cmb5 != "")
            {
                if (lln5 != "LIKE") likesign = "";

                mysql = mysql + " " + nor4 + " [" + cmb1 + "] " + lln5 + "'" + likesign + txt5 + likesign + "'";
            }

            if (cmb6 != "")
            {
                if (lln6 != "LIKE") likesign = "";

                mysql = mysql + " " + nor5 + " [" + cmb6 + "] " + lln6 + "'" + likesign + txt6 + likesign + "'";
            }

            if (cmb7 != "")
            {
                if (lln7 != "LIKE") likesign = "";

                mysql = mysql + " " + nor6 + " [" + cmb7 + "] " + lln7 + "'" + likesign + txt7 + likesign + "'";
            }

            if (cmb8 != "")
            {
                if (lln8 != "LIKE") likesign = "";

                mysql = mysql + " " + nor7 + " [" + cmb8 + "] " + lln8 + "'" + likesign + txt8 + likesign + "'";
            }

            if (cmb9 != "")
            {
                if (lln9 != "LIKE") likesign = "";

                mysql = mysql + " " + nor8 + " [" + cmb9 + "] " + lln9 + "'" + likesign + txt9 + likesign + "'";
            }

            if (cmb10 != "")
            {
                if (lln10 != "LIKE") likesign = "";

                mysql = mysql + " " + nor9 + " [" + cmb10 + "] " + lln10 + "'" + likesign + txt10 + likesign + "'";
            }

            if (cmb11 != "")
            {
                if (lln11 != "LIKE") likesign = "";

                mysql = mysql + " " + nor10 + " [" + cmb11 + "] " + lln11 + "'" + likesign + txt11 + likesign + "'";
            }

            if (cmb12 != "")
            {
                if (lln12 != "LIKE") likesign = "";

                mysql = mysql + " " + nor11 + " [" + cmb12 + "] " + lln12 + "'" + likesign + txt12 + likesign + "'";
            }

            if (cmb13 != "")
            {
                if (lln13 != "LIKE") likesign = "";

                mysql = mysql + " " + nor12 + " [" + cmb13 + "] " + lln13 + "'" + likesign + txt13 + likesign + "'";
            }

            if (cmb14 != "")
            {
                if (lln14 != "LIKE") likesign = "";

                mysql = mysql + " " + nor13 + " [" + cmb14 + "] " + lln14 + "'" + likesign + txt14 + likesign + "'";
            }

            if (cmb15 != "")
            {
                if (lln15 != "LIKE") likesign = "";

                mysql = mysql + " " + nor14 + " [" + cmb15 + "] " + lln15 + "'" + likesign + txt15 + likesign + "'";
            }

            if (cmb16 != "")
            {
                if (lln16 != "LIKE") likesign = "";

                mysql = mysql + " " + nor15 + " [" + cmb16 + "] " + lln16 + "'" + likesign + txt16 + likesign + "'";
            }


            label6.Text = mysql;

      
            var mnewsql = fullsql;

            if (fullsql.Contains("where"))
                mnewsql = mnewsql + " AND (" + mysql + ")";
            else
                mnewsql = mnewsql + " WHERE (" + mysql + ")";


            ClassPublicclass.Filllist(listView2, mnewsql);
            // shortsql = fullsql + " where " + mysql;
            filteredprint = mnewsql;
            //}
            
            label2.Text = listView2.Items.Count.ToString();

            // classPublicclass.generateFile(mnewsql);
        }

        private string filteredprint = "";

        private void Button5_Click_1(object sender, EventArgs e)
        {
            Adjustme();
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            panel6.Visible = false;
        }

          
      

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text == "ALL")
            {
                ClassMainclass.listselect = "";
            }
            else
            {
                ClassMainclass.listselect = "TOP "+ comboBox2.Text+" ";
            }
           // classPublicclass.generateFile(classMainclass.listselect);
            Filtermenow();
        }

      

        private void ComboBox2_KeyUp(object sender, KeyEventArgs e)
        {
            Filtermenow();
        }

        private void ClearFolder(string FolderName)
        {
            var dir = new DirectoryInfo(FolderName);

            foreach (var fi in dir.GetFiles())
            {
                fi.IsReadOnly = false;
                fi.Delete();
            }

            foreach (var di in dir.GetDirectories())
            {
                ClearFolder(di.FullName);
                di.Delete();
            }
        }

        private void FrmSystemMainList_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
               var path = ClassPublicclass.GetTemporaryDirectory("DE_Reports");
                ClearFolder(@path + "\\DE_Reports");
                path = ClassPublicclass.GetTemporaryDirectory("INTG_graphs");
                ClearFolder(@path + "\\INTG_graphs");
            }
            catch (Exception)
            {
                // classPublicclass.showMessage(n.Message.ToString()," Close all open files");
            }

        }

        private void Button7_Click(object sender, EventArgs e)
        {
            RefreshContent();
        }


        private void RefreshContent()
        {
            txtsimpleSearch.Text = "";
            ClassDatabaseConnection.dbTable = txtTableName.Text;

            ClassDatabaseConnection.cnn1 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + ClassDatabaseConnection.dbUserName + "; password =" +
                     ClassDatabaseConnection.dbPassword + "; Initial Catalog=" + ClassDatabaseConnection.dataBaseName + ";Data Source=" + ClassDatabaseConnection.serverName;
            Filtermenow();

            if (listView2.Columns.Count > 0)
            {
                advancedLinklabel.Visible = true;
            }
            else { advancedLinklabel.Visible = false; }
        }


        private void ListboxSQLServerInstances_DropDown(object sender, EventArgs e)
        {
            GetSQLDetails(listboxSQLServerInstances, "");
            listboxSQLServerInstances.Items.Add("(local)");
        }

        private void GetSQLDetails(ComboBox SQLcomboBox, string Sname)
        {
            var sie = new ClassSQLInfoEnumerator();
            try
            {
                if (Sname == "listboxSQLServerDatabaseInstances")
                {
                    SQLcomboBox.Items.Clear();
                    //sie.SQLServer = listboxSQLServerInstances.SelectedItem.ToString();
                    //sie.Username = textboxUserName.Text;
                    //sie.Password = textboxPassword.Text;
                    SQLcomboBox.Items.AddRange(sie.EnumerateSQLServersDatabases());
                }
                else
                {
                    SQLcomboBox.Items.Clear();
                    SQLcomboBox.Items.AddRange(sie.EnumerateSQLServers());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ListboxSQLServerInstances_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            { 
                txtServerName.Text = listboxSQLServerInstances.Text;

                ClassDatabaseConnection.serverName = listboxSQLServerInstances.Text;
                ClassDatabaseConnection.dataBaseName = txtDatabaseName.Text;
                ClassDatabaseConnection.dbUserName = textboxUserName.Text;
                ClassDatabaseConnection.dbPassword = textboxPassword.Text;
                ClassDatabaseConnection.dbTable = txtTableName.Text;


                
            }
            catch
            {
            }
        }

        private void ListboxSQLServerDatabaseInstances_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                txtDatabaseName.Text = listboxSQLServerDatabaseInstances.Text;




                ClassDatabaseConnection.serverName = listboxSQLServerInstances.Text;
                ClassDatabaseConnection.dataBaseName = txtDatabaseName.Text;
                ClassDatabaseConnection.dbUserName = textboxUserName.Text;
                ClassDatabaseConnection.dbPassword = textboxPassword.Text;
                ClassDatabaseConnection.dbTable = txtTableName.Text;
                 
            }
            catch
            {
            }
        }
        private bool SQLServerSelected()
        {
            if (listboxSQLServerInstances.SelectedIndex == -1)
                return false;
            else
                return true;
        }

        private bool UserDetailsEntered()
        {
            if (textboxUserName.Text != "" && textboxPassword.Text != "")
                return true;
            else
                return false;
        }

        private void ListboxSQLServerDatabaseInstances_DropDown(object sender, EventArgs e)
        {
            try
            {
                
                //show tables

                if (SQLServerSelected() || txtsimpleSearch.Text != "")
                    if (UserDetailsEntered())
                    {
                        var mynewcnn = "";
                        var dataBaseName = "";
                        mynewcnn = "Data Source=" + txtServerName.Text + ";Initial Catalog=" + dataBaseName + ";User ID=" +
                                   textboxUserName.Text + "; password =" + textboxPassword.Text +
                                   ";Integrated Security='" + checkBox2.Checked.ToString() + "'";
                        using (var con = new SqlConnection(mynewcnn))
                        {
                            con.Open();
                            // using (SqlCommand com = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", con))
                            using (var com = new SqlCommand("SELECT Name FROM sys.databases", con))
                            {
                                using (var reader = com.ExecuteReader())
                                {
                                    listboxSQLServerDatabaseInstances.Items.Clear();
                                    while (reader.Read()) listboxSQLServerDatabaseInstances.Items.Add((string)reader["Name"]);
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("A Username/Password Must Be Entered To View Database Information");
                    }
                else
                    MessageBox.Show("SQL Server Instance Must Be Selected To View Database Information");
            }
            catch (Exception le)
            {
                MessageBox.Show(le.Message);
                    
            }
        }

        private void CmbTableName_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtTableName.Text = cmbTableName.Text;

            ClassDatabaseConnection.serverName = listboxSQLServerInstances.Text;
            ClassDatabaseConnection.dataBaseName = txtDatabaseName.Text;
            ClassDatabaseConnection.dbUserName = textboxUserName.Text;
            ClassDatabaseConnection.dbPassword = textboxPassword.Text;
            ClassDatabaseConnection.dbTable = txtTableName.Text;
 
        }

        private void CmbTableName_DropDown(object sender, EventArgs e)
        {
          
          try
          {

              //show tables

              if (SQLServerSelected() || txtsimpleSearch.Text != "")
                  if (UserDetailsEntered())
                  {
                      var mynewcnn = "";
                     // var dataBaseName = "";
                      mynewcnn = "Data Source=" + txtServerName.Text + ";Initial Catalog=" + txtDatabaseName.Text + ";User ID=" +
                                 textboxUserName.Text + "; password =" + textboxPassword.Text +
                                 ";Integrated Security='" + checkBox2.Checked.ToString() + "'";
                      using (var con = new SqlConnection(mynewcnn))
                      {
                          con.Open();

                          var mysql = "SELECT TABLE_NAME FROM [" + txtDatabaseName.Text +
                           "].INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_NAME ASC";


                          // using (SqlCommand com = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", con))
                          using (var com = new SqlCommand(mysql, con))
                          {
                              using (var reader = com.ExecuteReader())
                              {
                                  cmbTableName.Items.Clear();
                                  while (reader.Read()) cmbTableName.Items.Add((string)reader["TABLE_NAME"]);
                              }
                          }
                      }
                  }
                  else
                  {
                      MessageBox.Show("A Username/Password Must Be Entered To View Database Information");
                  }
              else
                  MessageBox.Show("SQL Server Instance Must Be Selected To View Database Information");
          }
          catch (Exception le)
          {
              MessageBox.Show(le.Message);

          }
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            ClassPublicclass.GenerateFile(fullsql);

        }

     
    }
}