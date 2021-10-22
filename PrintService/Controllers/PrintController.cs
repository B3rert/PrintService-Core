using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PrintService.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Text.RegularExpressions;

using Spire.Pdf;
using System.Drawing.Printing;
using System.Management;
using PrintService.Utilities;

namespace PrintService.Controllers
{


    [Route("api/[controller]")]
    [ApiController]
    public class PrintController : ControllerBase
    {


        //Retorna 1, Sirve para verificar que el servicio funciona
        [HttpGet("connection")]
        public int getConnection()
        {
            return 1;
        }

        //Retorna una lista de imprsoras instaladas
        [HttpGet]
        public List<string> getPrinter()
        {

            var response = new List<string>();

            foreach (string strPrinter in PrinterSettings.InstalledPrinters)
            {
                response.Add(strPrinter);
            }
            return response;
        }
       
        
        //Genera un PDF con el texto recibido y lo imprime en la impreosra especificada
        [HttpPost("generate")]
        public IActionResult getPrint([FromBody] PrintModel print)
        {

            Globales.name_emited = print.name_emited;
            Globales.title_report = print.report_title;
            Globales.text_info = print.text_info;

            bool result = IsPrinterOnline(print.printer);
            if (!result)
            {
                return Ok(2);
            }

            var currentDirectory = Directory.GetCurrentDirectory(); //Ruta donden se encuntra el programa
            var absolutePath = $"{currentDirectory}\\testprinxt.pdf";
            var pagesPdf = 1;

            try
            {
                //Fuentes
                var color_red = new BaseColor(134, 13, 13);
                var color_gray = new BaseColor(127, 127, 127);
                var color_blue = new BaseColor(0, 0, 255);
                var normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);
                var normalFont9 = FontFactory.GetFont(FontFactory.HELVETICA, 9);
                var littleFont = FontFactory.GetFont(FontFactory.HELVETICA, 7);
                var littleFontBold = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7);
                var normalFontRed = FontFactory.GetFont(FontFactory.HELVETICA, 10, color_red);
                var boldFontRed = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, color_red);
                var boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                var boldFontRedHead = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10, color_red);
                var boldGray = FontFactory.GetFont(FontFactory.HELVETICA, 8, color_gray);
                var boldBlue = FontFactory.GetFont(FontFactory.HELVETICA, 7, color_blue);

                //Crear el documento
                Document doc = new Document(PageSize.LETTER, 20, 20,70,75);

                if (print.format == "Columnas")
                {
                   

                    PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(absolutePath, FileMode.Create));
                    writer.PageEvent = new HeaderFooter();
                    writer.AddViewerPreference(PdfName.PICKTRAYBYPDFSIZE, PdfBoolean.PDFTRUE);
                    //Head table body
                    Paragraph product_id = new Paragraph(print.column1, boldFontRedHead);
                    Paragraph product = new Paragraph(print.column2, boldFontRedHead);
                    Paragraph observation = new Paragraph(print.column3, boldFontRedHead);

                    doc.Open();

                    //Table body
                    PdfPTable body = new PdfPTable(3);
                    body.HeaderRows = 1;
                    body.AddCell(new PdfPCell(product_id) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0, PaddingBottom = 5, PaddingTop = 15 });
                    body.AddCell(new PdfPCell(product) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0, PaddingBottom = 5, PaddingTop = 15 });
                    body.AddCell(new PdfPCell(observation) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0, PaddingBottom = 5, PaddingTop = 15 });

                    //Datos de las demas filas
                    string[] rows = print.doc.Split("\n");
                    foreach (var item in rows)
                    {
                        string[] columns = item.Split(",,");
                        if (columns.Length == 2)
                        {

                            Paragraph product_id_value = new Paragraph("", normalFont);
                            Paragraph product_value = new Paragraph(columns[0], normalFont);
                            Paragraph observation_value = new Paragraph(columns[1], normalFont);

                            body.AddCell(returnCell(product_id_value));
                            body.AddCell(returnCell(product_value));
                            body.AddCell(returnCell(observation_value));

                        }
                        else
                        {
                            string text_value = "";

                            for (int i = 0; i < columns.Length; i++)
                            {
                                if (i != 0)
                                {
                                    text_value = text_value + columns[i];
                                }
                            }

                            Paragraph product_id_value = new Paragraph("", normalFont);
                            Paragraph product_value = new Paragraph(columns[0], normalFont);
                            Paragraph observation_value = new Paragraph(text_value, normalFont);

                            body.AddCell(returnCell(product_id_value));
                            body.AddCell(returnCell(product_value));
                            body.AddCell(returnCell(observation_value));
                        }
                    }

                    body.WidthPercentage = 100f;
                    body.PaddingTop = 20;
                    var colWidthPercentagesBody = new[] { 15f, 35f, 50f };
                    body.SetWidths(colWidthPercentagesBody);

                    doc.Add(body);

                    doc.Close();
                }
                else if (print.format == "Sin Columnas")
                {
                    PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(absolutePath, FileMode.Create));
                    writer.PageEvent = new HeaderFooter();
                    writer.AddViewerPreference(PdfName.PICKTRAYBYPDFSIZE, PdfBoolean.PDFTRUE);
                    //Head table body
                    Paragraph observation = new Paragraph(print.column3, boldFontRedHead);

                    doc.Open();

                    //Table body
                    PdfPTable body = new PdfPTable(1);
                    body.HeaderRows = 1;
                    body.AddCell(new PdfPCell(observation) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0, PaddingBottom = 5, PaddingTop = 15 });

                    //Datos de las demas filas
                    Paragraph observation_value = new Paragraph(print.doc, normalFont);

                    body.AddCell(returnCell(observation_value));

                    body.WidthPercentage = 100f;
                    body.PaddingTop = 20;
                    var colWidthPercentagesBody = new[] { 100f };
                    body.SetWidths(colWidthPercentagesBody);

                    doc.Add(body);

                    doc.Close();
                }
                else if (print.format == "Sin Formato")
                {
                    Document doc_none_format = new Document(PageSize.LETTER, 20, 20, 0, 0);

                    PdfWriter writer = PdfWriter.GetInstance(doc_none_format, new FileStream(absolutePath, FileMode.Create));
                    writer.AddViewerPreference(PdfName.PICKTRAYBYPDFSIZE, PdfBoolean.PDFTRUE);

                    // writer.PageEvent = new HeaderFooter();

                    Phrase conetnt = new Phrase();
                    conetnt.Add(new Chunk(print.name_emited, boldFont));
                    conetnt.Add(new Chunk($"\n\n{print.doc}", normalFont));


                    doc_none_format.Open();


                    doc_none_format.Add(conetnt);

                    doc_none_format.Close();
                }

               
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }


            //Obtiene el numero de páginas
            using (StreamReader sr = new StreamReader(System.IO.File.OpenRead(absolutePath)))
            {
                Regex regex = new Regex(@"/Type\s*//*Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());
                pagesPdf = matches.Count;
            }

            //si son mas de 10 páginas dividir el documento (Trabajarlo)

            //Configuracion de la impresion

            try
            {
                using (Spire.Pdf.PdfDocument doc = new Spire.Pdf.PdfDocument())
                {
                    doc.LoadFromFile(absolutePath);
                    doc.PrintSettings.PrinterName = print.printer;
                    doc.PrintSettings.Copies = (short)print.copies; //este metodo es mas rapido que hacer un for o foreach
                    doc.PrintSettings.SelectPageRange(1, pagesPdf);

                    doc.PrintSettings.PrintController = new StandardPrintController();

                    doc.Print();
                }
            }
            catch (Exception e)
            {

                return BadRequest(e.Message);
            }

            //Elimina el PDF que usó para la impresion
            if (System.IO.File.Exists(absolutePath))
            {
                while (System.IO.File.Exists(absolutePath))
                {
                    System.IO.File.Delete(absolutePath);
                }

            }

            return Ok(1);
        }

        private PdfPCell returnCell(Paragraph text)
        {
            PdfPCell cell = new PdfPCell();
            cell.AddElement(text);
            cell.BorderWidthBottom = 0;
            cell.BorderWidthLeft = 0;
            cell.BorderWidthTop = 0;
            cell.BorderWidthRight = 0;
            cell.Padding = 0;

            return cell;
        }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validar la compatibilidad de la plataforma", Justification = "<pendiente>")]
        private bool IsPrinterOnline(string printerName)
        {
            string str = "";
            bool online = false;

            //set the scope of this search to the local machine
            ManagementScope scope = new ManagementScope(ManagementPath.DefaultPath);
            //connect to the machine
            scope.Connect();

            //query for the ManagementObjectSearcher
            SelectQuery query = new SelectQuery("select * from Win32_Printer");

            ManagementClass m = new ManagementClass("Win32_Printer");

            ManagementObjectSearcher obj = new ManagementObjectSearcher(scope, query);

            //get each instance from the ManagementObjectSearcher object
            using (ManagementObjectCollection printers = m.GetInstances())
                //now loop through each printer instance returned
                foreach (ManagementObject printer in printers)
                {
                    //first make sure we got something back
                    if (printer != null)
                    {
                        //get the current printer name in the loop
                        str = printer["Name"].ToString().ToLower();

                        //check if it matches the name provided
                        if (str.Equals(printerName.ToLower()))
                        {
                            //since we found a match check it's status
                            if (printer["WorkOffline"].ToString().ToLower().Equals("true") || printer["PrinterStatus"].Equals(7))
                                //it's offline
                                online = false;
                            else
                                //it's online
                                online = true;
                        }
                    }
                    else
                        throw new Exception("No printers were found");
                }
            return online;
        }
    }
}
