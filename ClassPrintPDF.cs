using System;
using System.Text;
using System.Data.OleDb;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Windows.Forms;
using System.Drawing;

namespace DataExplorer2
{
    class ClassPrintPDF
    {
        

        /// <summary>
        /// MAKE PDF DOCUMENT COLORFUL
        /// </summary>
        public static iTextSharp.text.BaseColor myborderColor = new BaseColor(Color.White);
        public static iTextSharp.text.BaseColor mybaseColor = new BaseColor(Color.Lime);
        public static iTextSharp.text.BaseColor mybaseFontColor = new BaseColor(Color.Black);
        public static iTextSharp.text.BaseColor mybaseColora = new BaseColor(Color.Lime);
        public static iTextSharp.text.BaseColor mybaseFontColora = new BaseColor(Color.Black);
        public static iTextSharp.text.BaseColor mytitleColor = new BaseColor(Color.DarkGreen);
        public static iTextSharp.text.BaseColor mytitleFontColor = new BaseColor(Color.White);
        public static iTextSharp.text.BaseColor myheaderColor = new BaseColor(Color.DarkGreen);
        public static iTextSharp.text.BaseColor myheaderFontColor = new BaseColor(Color.White);
        public static iTextSharp.text.BaseColor myfooterColor = new BaseColor(Color.DarkGreen);
        public static iTextSharp.text.BaseColor myfooterFontColor = new BaseColor(Color.White);
        public static System.Drawing.Color myListColor = new System.Drawing.Color();
        public static System.Drawing.Color myListfontColor = new System.Drawing.Color();
        public static System.Drawing.Color myListColora = new System.Drawing.Color();
        public static System.Drawing.Color myListfontColora = new System.Drawing.Color();
    //    public static String path = System.Windows.Forms.Application.StartupPath.ToString();



        public static void Generatereportpdf1(string Title, string extension, string sql, string orientation)
        {
           
            
              var  path = ClassPublicclass.GetTemporaryDirectory("DE_Reports");            

                Random rnd = new Random();
                int surfix = rnd.Next(52);
                string sFilePDF = Title + surfix.ToString();
                string Reports = "DE_Reports";
            
                try
                {
                    iTextSharp.text.Font font = new iTextSharp.text.Font();
                    iTextSharp.text.Font boldfont = new iTextSharp.text.Font();
                    iTextSharp.text.Font fontclear = new iTextSharp.text.Font();
                    iTextSharp.text.Font fontfooter = new iTextSharp.text.Font();
                    fontfooter = FontFactory.GetFont(FontFactory.COURIER, 7, iTextSharp.text.Font.NORMAL);
                    iTextSharp.text.Font fontfooters = new iTextSharp.text.Font();
                    fontfooters = FontFactory.GetFont(FontFactory.COURIER, 7, iTextSharp.text.Font.NORMAL);
                    font = FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.NORMAL);
                    boldfont = FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.BOLD);
                    fontclear = FontFactory.GetFont(FontFactory.HELVETICA, 7, iTextSharp.text.Font.ITALIC);
                    iTextSharp.text.Document document = new iTextSharp.text.Document(PageSize.A4, 50, 50, 50, 50);
                    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(String.Format("{0}\\{1}\\{2}.{3}", path, Reports, sFilePDF, extension), FileMode.Create));
                    writer.PageEvent = new ClassPDFFooter();

                    if (orientation == "Landscape" || orientation == "L")
                    {
                        document.SetPageSize(new iTextSharp.text.Rectangle(792, 612)); //landscape
                        ClassPDFFooter.fwidth = 692F;
                    }
                    else
                    {
                        ClassPDFFooter.fwidth = 495F;
                    }
                    // step 3: we open the document
                    document.Open();
                PdfPTable titleTable = new PdfPTable(10)
                {
                    WidthPercentage = 100
                };//10 column
                PdfPCell cell = new PdfPCell(new Phrase("Title ", new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 24F)));
                    Phrase mytitle = new Phrase(Title, boldfont);
                    boldfont.Color = mytitleFontColor;
                    cell = new PdfPCell();
                    cell.AddElement(mytitle);
                    cell.Colspan = 8;
                    cell.Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER;
                    cell.BackgroundColor = mytitleColor;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    titleTable.AddCell(cell);
                    Phrase printdate = new Phrase("Date " + System.DateTime.Now.Date.ToString(ClassMainclass.Dateformat) + " " + System.DateTime.Now.ToShortTimeString(), fontfooters);
                    fontfooters.Color = mytitleFontColor;
                    cell = new PdfPCell();
                    cell.AddElement(printdate);
                    cell.Colspan = 2;
                    cell.Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER;
                    cell.BackgroundColor = mytitleColor;
                    //  printdate.Font.Color = mytitleFontColor;
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                    titleTable.AddCell(cell);
                    PdfPTable spacetable = new PdfPTable(1);//1,1
                    Phrase spacecontent = new Phrase(" ", boldfont);
                    spacecontent.Font.Color = iTextSharp.text.BaseColor.BLACK;
                cell = new PdfPCell(new Phrase(spacecontent))
                {
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    Border = iTextSharp.text.Rectangle.NO_BORDER | iTextSharp.text.Rectangle.NO_BORDER
                };
                spacetable.AddCell(cell);
                    Cursor.Current = Cursors.WaitCursor;
                    StringBuilder sbConnCOMBOx = new StringBuilder();
                    sbConnCOMBOx.Append(ClassDatabaseConnection.cnn1);
                    sbConnCOMBOx.Append(";Extended Properties=");
                    sbConnCOMBOx.Append(Convert.ToChar(34));
                    sbConnCOMBOx.Append(Convert.ToChar(34));
                    OleDbConnection cnExcelCOMBOx = new OleDbConnection(sbConnCOMBOx.ToString());
                    cnExcelCOMBOx.Open();
                    StringBuilder sbSQLCOMBOx = new StringBuilder();
                    sbSQLCOMBOx.Append(sql);
                    OleDbCommand cmdExcelCOMBOx = new OleDbCommand(sbSQLCOMBOx.ToString(), cnExcelCOMBOx);
                    OleDbDataReader drExcelCOMBOx = cmdExcelCOMBOx.ExecuteReader();

                    int nocols = 10;
                    int mycolspan = 1;

                    if (drExcelCOMBOx.FieldCount + 1 == 2) { nocols = 10; mycolspan = 9; }
                    else if (drExcelCOMBOx.FieldCount + 1 == 3) { nocols = 11; mycolspan = 5; }
                    else if (drExcelCOMBOx.FieldCount + 1 == 4) { nocols = 10; mycolspan = 3; }
                    else if (drExcelCOMBOx.FieldCount + 1 == 5) { nocols = 13; mycolspan = 3; }
                    else if (drExcelCOMBOx.FieldCount + 1 == 6) { nocols = 11; mycolspan = 2; }
                    else if (drExcelCOMBOx.FieldCount + 1 == 7) { nocols = 13; mycolspan = 2; }
                    else if (drExcelCOMBOx.FieldCount + 1 == 8) { nocols = 15; mycolspan = 2; }
                    else if (drExcelCOMBOx.FieldCount + 1 == 9) { nocols = 17; mycolspan = 2; }
                    else if (drExcelCOMBOx.FieldCount + 1 == 10) { nocols = 10; mycolspan = 1; }
                    else { nocols = drExcelCOMBOx.FieldCount + 1; }
                    PdfPTable contentTable = new PdfPTable(nocols);

                    //PdfPTable TotalsTable = new PdfPTable(nocols);
                    Phrase indextitle = new Phrase("No", boldfont);
                    indextitle.Font.Color = myheaderFontColor;
                cell = new PdfPCell(new Phrase(indextitle))
                {
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    BackgroundColor = myheaderColor,
                    Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE,
                    BorderColor = myborderColor
                };
                contentTable.AddCell(cell);

                    Phrase totaltitle = new Phrase("Totals", boldfont);
                    totaltitle.Font.Color = myheaderFontColor;
                cell = new PdfPCell(new Phrase(totaltitle))
                {
                    HorizontalAlignment = Element.ALIGN_LEFT,
                    BackgroundColor = myheaderColor,
                    Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE,
                    BorderColor = myborderColor
                };
                // TotalsTable.AddCell(cell);


                // double tots1 = 0.00;
                for (int i = 0; i < drExcelCOMBOx.FieldCount; i++)
                    {
                        if (drExcelCOMBOx.GetName(i).ToUpper() == "ID")
                        {
                            Phrase content = new Phrase("No", boldfont);
                            content.Font.Color = myheaderFontColor;
                        cell = new PdfPCell(new Phrase(content))
                        {
                            HorizontalAlignment = Element.ALIGN_LEFT,
                            BackgroundColor = myheaderColor,
                            Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE,
                            BorderColor = myborderColor,
                            Colspan = mycolspan
                        };
                        contentTable.AddCell(cell);
                            //TotalsTable.AddCell(cell);
                        }

                        else if (drExcelCOMBOx.GetName(i).ToUpper() == "DESC")
                        {
                            Phrase content = new Phrase("Description", boldfont);
                            content.Font.Color = myheaderFontColor;
                        cell = new PdfPCell(new Phrase(content))
                        {
                            HorizontalAlignment = Element.ALIGN_LEFT,
                            BackgroundColor = myheaderColor,
                            Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE,
                            BorderColor = myborderColor,
                            Colspan = mycolspan
                        };
                        contentTable.AddCell(cell);
                        }
                        else if (drExcelCOMBOx.GetName(i).ToUpper() == "% OF" || drExcelCOMBOx.GetName(i).ToUpper() == "% AMOUNT" || drExcelCOMBOx.GetName(i).ToUpper() == "% RATE" || drExcelCOMBOx.GetName(i).ToUpper() == "DAY" || drExcelCOMBOx.GetName(i).ToUpper() == "MONTH" || drExcelCOMBOx.GetName(i).ToUpper() == "YEAR" || drExcelCOMBOx.GetName(i).ToUpper() == "MONTH" || drExcelCOMBOx.GetName(i).ToUpper() == "SENDER" || drExcelCOMBOx.GetName(i).ToUpper() == "TOTAL DEDUCTIONS" || drExcelCOMBOx.GetName(i).ToUpper() == "PENSION" || drExcelCOMBOx.GetName(i).ToUpper() == "NHIF" || drExcelCOMBOx.GetName(i).ToUpper() == "NSSF" || drExcelCOMBOx.GetName(i).ToUpper() == "MPR" || drExcelCOMBOx.GetName(i).ToUpper() == "TAX CHARGED" || drExcelCOMBOx.GetName(i).ToUpper() == "NET PAY" || drExcelCOMBOx.GetName(i).ToUpper() == "GROSS PAY" || drExcelCOMBOx.GetName(i).ToUpper() == "BASIC PAY" || drExcelCOMBOx.GetName(i).ToUpper() == "Taxable Pay" || drExcelCOMBOx.GetName(i).ToUpper() == "PAYE" || drExcelCOMBOx.GetName(i).ToUpper() == "AMOUNT" || drExcelCOMBOx.GetName(i).ToUpper() == "QTY" || drExcelCOMBOx.GetName(i).ToUpper() == "ORDER NO" || drExcelCOMBOx.GetName(i).ToUpper() == "PRICE" || drExcelCOMBOx.GetName(i).ToUpper() == "VAT")
                        {
                            Phrase cnt = new Phrase(drExcelCOMBOx.GetName(i).ToString(), boldfont);
                        Paragraph amnt = new Paragraph(cnt)
                        {
                            Alignment = Element.ALIGN_RIGHT
                        };
                        amnt.Font.Color = myheaderFontColor;
                            cnt.Font.Color = myheaderFontColor;
                        cell = new PdfPCell(new Phrase(amnt))
                        {
                            HorizontalAlignment = Element.ALIGN_RIGHT,
                            BackgroundColor = myheaderColor,
                            Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE,
                            BorderColor = myborderColor,
                            Colspan = mycolspan
                        };
                        contentTable.AddCell(cell);

                        }

                        else if (drExcelCOMBOx.GetName(i).ToUpper() == "Usergroup")
                        {
                            Phrase content = new Phrase("User Group", boldfont);
                            content.Font.Color = myheaderFontColor;
                        cell = new PdfPCell(new Phrase(content))
                        {
                            HorizontalAlignment = Element.ALIGN_LEFT,
                            BackgroundColor = myheaderColor,
                            Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE,
                            BorderColor = myborderColor,
                            Colspan = mycolspan
                        };
                        contentTable.AddCell(cell);
                            //    TotalsTable.AddCell(cell);
                        }

                        else if (drExcelCOMBOx.GetName(i).ToUpper() == "DEFINITION")
                        {
                            Phrase content = new Phrase(Title, boldfont);
                            content.Font.Color = myheaderFontColor;
                        cell = new PdfPCell(new Phrase(content))
                        {
                            HorizontalAlignment = Element.ALIGN_LEFT,
                            BackgroundColor = myheaderColor,
                            Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE,
                            BorderColor = myborderColor,
                            Colspan = mycolspan
                        };
                        contentTable.AddCell(cell);
                            //   TotalsTable.AddCell(cell);
                        }

                        else
                        {
                            Phrase content = new Phrase(drExcelCOMBOx.GetName(i).Replace("_", " "), boldfont);
                            content.Font.Color = myheaderFontColor;
                        cell = new PdfPCell(new Phrase(content))
                        {
                            HorizontalAlignment = Element.ALIGN_LEFT,
                            BackgroundColor = myheaderColor,
                            Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE,
                            BorderColor = myborderColor,
                            Colspan = mycolspan
                        };
                        contentTable.AddCell(cell);
                            //   TotalsTable.AddCell(cell);
                        }
                    }
                    cnExcelCOMBOx.Close();
                    StringBuilder sbConnCOMBO = new StringBuilder();
                    sbConnCOMBO.Append(ClassDatabaseConnection.cnn1);
                    sbConnCOMBO.Append(";Extended Properties=");
                    sbConnCOMBO.Append(Convert.ToChar(34));
                    sbConnCOMBO.Append(Convert.ToChar(34));
                    OleDbConnection cnExcelCOMBO = new OleDbConnection(sbConnCOMBO.ToString());
                    cnExcelCOMBO.Open();
                    StringBuilder sbSQLCOMBO = new StringBuilder();
                    sbSQLCOMBO.Append(sql);
                    OleDbCommand cmdExcelCOMBO = new OleDbCommand(sbSQLCOMBO.ToString(), cnExcelCOMBO);
                    OleDbDataReader drExcelCOMBO = cmdExcelCOMBO.ExecuteReader();
                    int j = 1;
                    while (drExcelCOMBO.Read())
                    {
                        Phrase index = new Phrase(j.ToString() + ".", font);

                    cell = new PdfPCell(new Phrase(index))
                    {
                        HorizontalAlignment = Element.ALIGN_LEFT,
                        BackgroundColor = iTextSharp.text.BaseColor.WHITE
                    };

                    if (j % 2 != 0)
                        {
                            cell.BackgroundColor = mybaseColora;
                            index.Font.Color = mybaseFontColora;
                        }
                        else
                        {
                            cell.BackgroundColor = mybaseColor;
                            index.Font.Color = mybaseFontColor;
                        }

                        cell.Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE;
                        cell.BorderColor = myborderColor;
                        contentTable.AddCell(cell);

                        for (int K = 0; K < drExcelCOMBO.FieldCount; K++)
                        {
                            Phrase content = new Phrase(drExcelCOMBO[drExcelCOMBO.GetName(K)].ToString().Replace('_', ' '), font);
                            try
                            {
                                if ((double.Parse(drExcelCOMBO[drExcelCOMBO.GetName(K)].ToString()) <= 0 || double.Parse(drExcelCOMBO[drExcelCOMBO.GetName(K)].ToString()) >= 0) && drExcelCOMBO.GetName(K).ToString() != "Year" && drExcelCOMBO.GetName(K).ToString() != "Month Period" && drExcelCOMBO.GetName(K).ToString() != "Year Period" && drExcelCOMBO.GetName(K).ToString() != "NHIF No" && drExcelCOMBO.GetName(K).ToString() != "NSSF NO" && drExcelCOMBO.GetName(K).ToString() != "Month" && drExcelCOMBO.GetName(K).ToString() != "Day" && drExcelCOMBO.GetName(K).ToString() != "Payroll No" && drExcelCOMBO.GetName(K).ToString() != "ID" && drExcelCOMBO.GetName(K).ToString() != "Account No" && drExcelCOMBO.GetName(K).ToString() != "Total From" && drExcelCOMBO.GetName(K).ToString() != "Total To")
                                {
                                    // if (drExcelCOMBO.GetName(K).ToString() != "Year" || drExcelCOMBO.GetName(K).ToString() != "Day" || drExcelCOMBO.GetName(K).ToString() != "Month" || drExcelCOMBO.GetName(K).ToString() != "ID")
                                    //  {
                                    content = new Phrase(string.Format("{0:0.00}", double.Parse(drExcelCOMBO[drExcelCOMBO.GetName(K)].ToString())), font);
                                }

                                else
                                {
                                    content = new Phrase((drExcelCOMBO[drExcelCOMBO.GetName(K)].ToString()), font);
                                }
                                //   }
                            }
                            catch (Exception) { }
                            Paragraph nn = new Paragraph(content);
                            Label getcolname = new Label();

                            try
                            {
                                if ((double.Parse(drExcelCOMBO[drExcelCOMBO.GetName(K)].ToString()) <= 0 || double.Parse(drExcelCOMBO[drExcelCOMBO.GetName(K)].ToString()) >= 0) && drExcelCOMBO.GetName(K) != "Month" || drExcelCOMBO.GetName(K) != "Year" && drExcelCOMBO.GetName(K) != "Month Period" || drExcelCOMBO.GetName(K) != "Year Period" || drExcelCOMBO.GetName(K) != "NHIF No" || drExcelCOMBO.GetName(K) != "NSSF NO" || drExcelCOMBO.GetName(K) != "Day" || drExcelCOMBO.GetName(K) != "Payroll No")
                                {
                                    nn.Alignment = Element.ALIGN_RIGHT;
                                }
                                else
                                {
                                    nn.Alignment = Element.ALIGN_LEFT;
                                }
                            }
                            catch (Exception) { }

                            cell = new PdfPCell(new Phrase(nn));

                            try
                            {
                                if (double.Parse(drExcelCOMBO[drExcelCOMBO.GetName(K)].ToString()) >= 0)
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_RIGHT;
                                }
                                else
                                {
                                    cell.HorizontalAlignment = Element.ALIGN_LEFT;
                                }
                            }
                            catch (Exception) { }

                            if (j % 2 != 0)
                            {
                                cell.BackgroundColor = mybaseColora;
                                index.Font.Color = mybaseFontColora;
                            }
                            else
                            {
                                cell.BackgroundColor = mybaseColor;
                                content.Font.Color = mybaseFontColor;
                            }
                            cell.Border = iTextSharp.text.Rectangle.TOP_BORDER | iTextSharp.text.Rectangle.RECTANGLE;
                            cell.BorderColor = myborderColor;
                            cell.Colspan = mycolspan;
                            contentTable.AddCell(cell);
                        }
                        j++;
                    }
                    cnExcelCOMBO.Close();
                    contentTable.WidthPercentage = 100;
                    // TotalsTable.WidthPercentage = 100;
                    document.Add(spacetable);
                    document.Add(spacetable);
                    document.Add(spacetable);
                    document.Add(spacetable);
                    document.Add(spacetable);
                    document.Add(spacetable);
                    document.Add(titleTable);
                    document.Add(spacetable);
                    document.Add(contentTable);
                    document.Add(spacetable);
                    document.Add(spacetable);
                    // document.Add(TotalsTable);
                    document.Close();
                    try
                    {
                        System.Diagnostics.Process.Start(String.Format("{0}\\{1}\\{2}.{3}", path, Reports, sFilePDF, extension));
                    }
                    catch (Exception)
                    {
                        // classPublicclass.showMessage("The Pdf report may be open, please close it and try again");
                    }
                }
                catch (Exception n)
                {
                    MessageBox.Show("Document may be open please, close it and try again " + n.ToString());
                   
                }

            }
        }
          
    
    }
        

