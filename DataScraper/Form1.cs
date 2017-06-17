using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataScraper
{
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            
            InitializeComponent();
        }

        

        private void clickMe_Click(object sender, EventArgs e)
        {
            DateTime date = DateTime.Today;
            MessageBox.Show("Button Clicked");

            string Url = "https://www.eziline.com/";
            HtmlWeb web = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = web.Load(Url);

            string Metascore = doc.DocumentNode.SelectNodes("//*[@id=\"header\"]/div[1]/div/div/div/div[1]/h2")[0].InnerText;
            string Title= doc.DocumentNode.SelectNodes("/html/head/title")[0].InnerText;

            Microsoft.Office.Interop.Excel.Application oXL;
            Microsoft.Office.Interop.Excel._Workbook oWB;
            Microsoft.Office.Interop.Excel._Worksheet oSheet;
            

            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;

            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oSheet.Cells[1, 1] = "Title";
            oSheet.Cells[1, 2] = "Data";
            oSheet.Cells[1, 3] = "Date";
            

            //Format A1:D1 as bold, vertical alignment = center.
            oSheet.get_Range("A1", "D1").Font.Bold = true;
            oSheet.get_Range("A1", "D1").VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Create an array to multiple values at once.
            string[,] saNames = new string[5, 3];

            saNames[0, 0] = Title;
            saNames[0, 1] = Metascore;
            saNames[0, 2]= date.ToString("dd/MM/yyyy");
            
            saNames[1, 2] = date.ToString("dd/MM/yyyy");

            //Fill A2:B6 with an array of values (First and Last Names).
            oSheet.get_Range("A2", "C6").Value2 = saNames;
            
            oWB.SaveAs("test505.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            
            
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string temp = "DOne";
            MessageBox.Show(temp);
        }
    }
   
}
