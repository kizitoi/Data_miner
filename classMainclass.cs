using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataExplorer2
{
    public static class ClassMainclass
    {
        /////THEME COLORS

        public static Color ToolstripColor = Color.CornflowerBlue;
        public static Color TabColor = Color.CornflowerBlue;
        public static Color ListHeaderColor = Color.CornflowerBlue;
        public static Color ToolstripFontColor = Color.White;
        public static Color FormColor = Color.White;


        /// <summary>
        /// ////////////////
        /// </summary>

        public static string Dateformat = "yyyy-MM-dd"; // specify the date format
        public static string focussearchitem = "";
        public static string listselect = "TOP 1000"; // SELECT TOP 1000 ITEMS FOR THE LIST
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
           
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmSystemMainList());
        }
    }
}
