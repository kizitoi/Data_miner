using System;
using System.Drawing;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace DataExplorer2
{
    public class ClassPDFFooter : PdfPageEventHelper
    {
        public static float fwidth = 1F;
        private string path = Application.StartupPath.ToString();
        private Label companyNameLbl = new Label();
        private Label comptellbl = new Label();
        private Label mobilelbl = new Label();
        private Label compaddlbl = new Label();
        private readonly Label companylogo1lbl = new Label();
        private Label websitelbl = new Label();
        private Label emaillbl = new Label();
        private Label faxlbl = new Label();

        private Label street = new Label();

        // write on top of document
        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            var font = new iTextSharp.text.Font();
            var boldfont = new iTextSharp.text.Font();
            var fontclear = new iTextSharp.text.Font();
            var fontfooter = new iTextSharp.text.Font();
            font = FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL);
            boldfont = FontFactory.GetFont(FontFactory.HELVETICA, 9, iTextSharp.text.Font.BOLD);
            fontclear = FontFactory.GetFont(FontFactory.HELVETICA, 9, iTextSharp.text.Font.ITALIC);
            fontfooter = FontFactory.GetFont(FontFactory.COURIER, 5, iTextSharp.text.Font.NORMAL);

            companyNameLbl.Text = "Company Name";
 
            path = ClassPublicclass.GetTemporaryDirectory("INTG_Images");

            iTextSharp.text.Image companylogo;

            try
            {
                companylogo = iTextSharp.text.Image.GetInstance(new Uri(path + "\\INTG_Images\\clogo.bmp"));
            }
            catch (Exception)
            {
                // companylogo = iTextSharp.text.Image.GetInstance(new Uri(Application.StartupPath + "\\Images\\clogo.bmp"));
                companylogo = null;
            }


            //classPublicclass.showMessage(companylogo.ToString());
            var cname = new Phrase(companyNameLbl.Text, font);
            cname.Font.Color = BaseColor.BLACK;

            var cadd = new Phrase(compaddlbl.Text, font);
            cadd.Font.Color = BaseColor.BLACK;
            var cstreet = new Phrase(street.Text, font);
            cstreet.Font.Color = BaseColor.BLACK;

            var aTable = new PdfPTable(4)
            {
                WidthPercentage = 100
            }; //4 columns

            var cell = new PdfPCell(new Phrase("Header",
                new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 24F)));

            if (companylogo != null)
            {
                cell = new PdfPCell();
                companylogo.ScaleAbsolute(80f, 80f);
                cell.AddElement(companylogo);
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER;
                aTable.AddCell(cell);
            }
            else
            {
                cell = new PdfPCell();
                var nph = new Phrase("", font);
                cell.AddElement(nph);
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER;
                aTable.AddCell(cell);
            }

            cell = new PdfPCell
            {
                HorizontalAlignment = Element.ALIGN_LEFT
            };
            cell.AddElement(cname);
            cell.AddElement(cstreet);
            cell.AddElement(cadd);
            cell.Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER;
            aTable.AddCell(cell);


            var ctel = new Phrase("Tel: " + comptellbl.Text, font);
            var cmob = new Phrase("Cell: " + mobilelbl.Text, font);
            var cfax = new Phrase("Fax: " + faxlbl.Text, font);

            cell = new PdfPCell
            {
                HorizontalAlignment = Element.ALIGN_LEFT
            };

            if (comptellbl.Text != "") cell.AddElement(ctel);

            if (mobilelbl.Text != "") cell.AddElement(cmob);

            if (faxlbl.Text != "") cell.AddElement(cfax);

            cell.Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER;
            aTable.AddCell(cell);

            var cemail = new Phrase("Email: " + emaillbl.Text, font);
            var cweb = new Phrase("Web: " + websitelbl.Text, font);

            ctel.Font.Color = BaseColor.BLACK;
            cell = new PdfPCell();

            if (emaillbl.Text != "") cell.AddElement(cemail);

            if (websitelbl.Text != "") cell.AddElement(cweb);

            // cell.HorizontalAlignment = Element.ALIGN_JUSTIFIED_ALL;
            cell.Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER;
            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            aTable.AddCell(cell);
            aTable.WidthPercentage = 100;

            //base.OnOpenDocument(writer, document);
            base.OnStartPage(writer, document);
            var header = new PdfPTable(new float[] { 1F })
            {
                SpacingAfter = 10F,
                ///PdfPCell cell;
                TotalWidth = fwidth
            };
            cell = new PdfPCell
            {
                Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER
            };
            cell.AddElement(aTable);
            header.AddCell(cell);
            header.WriteSelectedRows(0, -1, 50, document.Top, writer.DirectContent);
        }

        // write on start of each page
        public override void OnStartPage(PdfWriter writer, Document document)
        {
            base.OnStartPage(writer, document);
        }

        // write on end of each page


        private readonly BaseColor myfooterColor = new BaseColor(Color.DarkGreen);
        private readonly BaseColor myfooterFontColor = new BaseColor(Color.White);

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            var font = new iTextSharp.text.Font();
            var boldfont = new iTextSharp.text.Font();
            var fontclear = new iTextSharp.text.Font();
            var fontfooter = new iTextSharp.text.Font();
            font = FontFactory.GetFont(FontFactory.HELVETICA, 5, iTextSharp.text.Font.NORMAL);
            boldfont = FontFactory.GetFont(FontFactory.HELVETICA, 9, iTextSharp.text.Font.BOLD);
            fontclear = FontFactory.GetFont(FontFactory.HELVETICA, 9, iTextSharp.text.Font.ITALIC);
            fontfooter = FontFactory.GetFont(FontFactory.COURIER, 5, iTextSharp.text.Font.NORMAL, myfooterFontColor);

            companyNameLbl.Text = "Company Name";
 
 

            var myfooter =
                new Phrase(
                    companyNameLbl.Text + " " , fontfooter);
            var mypage = new Phrase("Page: " + document.PageNumber.ToString(), fontfooter);
            base.OnEndPage(writer, document);
            var footerTable = new PdfPTable(12);
            //PdfPTable footerTable = new PdfPTable(new float[] { 1F });
            PdfPCell cell;
            footerTable.TotalWidth = fwidth;
            //footerTable.WidthPercentage = 100;            

            cell = new PdfPCell(myfooter)
            {
                BackgroundColor = myfooterColor,
                Colspan = 11,
                Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER
            };
            footerTable.AddCell(cell);

            /*cell = new PdfPCell();
            cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;            
            cell.Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER;
            footerTable.AddCell(cell);*/

            cell = new PdfPCell();
            cell.AddElement(mypage);
            cell.BackgroundColor = myfooterColor;
            cell.HorizontalAlignment = Element.ALIGN_RIGHT;
            cell.VerticalAlignment = Element.ALIGN_TOP;
            cell.Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER;
            footerTable.AddCell(cell);
            footerTable.WriteSelectedRows(0, -1, 50, document.Bottom, writer.DirectContent);
        }

        //write on close of document
        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);
        }
    }
}