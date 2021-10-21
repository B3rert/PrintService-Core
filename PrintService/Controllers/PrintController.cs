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

namespace PrintService.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PrintController : ControllerBase
    {
        public readonly string p_name_emited;

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
       
        class HeaderFooter : PdfPageEventHelper
        {

            
            public override void OnEndPage(PdfWriter writer, Document document)
            {

                var currentDirectory = Directory.GetCurrentDirectory(); //Ruta donden se encuntra el programa

                PdfContentByte cb = writer.DirectContent;
                
                
                //Fuentes
                var color_red = new BaseColor(134, 13, 13);
                var color_gray = new BaseColor(127, 127, 127);
                var color_blue = new BaseColor(0, 0, 255);
                var normalFont9 = FontFactory.GetFont(FontFactory.HELVETICA, 9);
                var boldFontRed = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, color_red);
                var boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                var boldGray = FontFactory.GetFont(FontFactory.HELVETICA, 8, color_gray);
                var boldBlue = FontFactory.GetFont(FontFactory.HELVETICA, 7, color_blue);

                //Logos
                Image logo_empresa = Image.GetInstance($"{currentDirectory}\\assets\\empresa_logo.jfif");
                logo_empresa.ScalePercent(50f);
                logo_empresa.SetAbsolutePosition(465, 705);
                Image logo_dev = Image.GetInstance($"{currentDirectory}\\assets\\demosoft.jfif");
                logo_dev.ScalePercent(10f);
                logo_dev.SetAbsolutePosition(535, 15);

                cb.AddImage(logo_empresa);
                cb.AddImage(logo_dev);

                //Header
                Paragraph title_report = new Paragraph("Receta", boldFontRed);
                Paragraph name_emmited = new Paragraph("Dr. Cabrera", boldFont);

                PdfPTable header = new PdfPTable(1);
                header.TotalWidth = document.Right - document.Left;

                header.AddCell(new PdfPCell(title_report) { HorizontalAlignment = Element.ALIGN_CENTER, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                header.AddCell(new PdfPCell(new Paragraph(" ")) {  BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                header.AddCell(new PdfPCell(new Paragraph(" ")) {  BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                header.AddCell(new PdfPCell(name_emmited) { HorizontalAlignment = Element.ALIGN_LEFT,  BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });

                header.WriteSelectedRows(0, -1, 20, 790, writer.DirectContent);

                //Footer
                //DateTime
                DateTime dateTime = DateTime.Now;
                var hora = dateTime.Hour;
                var suffix = (hora >= 12) ? " p.m." : " a.m.";
                string fecha_hora = dateTime.ToString("dd/MM/yyyy hh:mm:ss") + suffix;

                Phrase paginator = new Phrase();
                paginator.Add(new Chunk(fecha_hora, boldGray));
                paginator.Add(new Chunk($" Página {document.PageNumber}", normalFont9));
                Paragraph text_info = new Paragraph("PBX: 2259-3232 / 6a. Ave \"A\" 13-25 Zona 9, Guatemala/ info@imcguate.com\ndrcabreramancio@imcguate.com    5552-417    /IMCCabreraMancio/\nwww.imcguate.com", boldBlue);

                PdfPTable footer = new PdfPTable(3);
                footer.TotalWidth = document.Right - document.Left;

                var colWidthPercentagesFooter = new[] { 40f, 50f, 10f };
                footer.SetWidths(colWidthPercentagesFooter);

                footer.AddCell(new PdfPCell(paginator) { HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_BOTTOM, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                footer.AddCell(new PdfPCell(text_info) { HorizontalAlignment = Element.ALIGN_CENTER, VerticalAlignment = Element.ALIGN_CENTER, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                footer.AddCell(new PdfPCell(new Paragraph(" ")) {  BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });

                footer.WriteSelectedRows(0, -1, 20,40, writer.DirectContent);

            }
        }

        
        //Genera un PDF con el texto recibido y lo imprime en la impreosra especificada
        [HttpPost("generate")]
        public IActionResult getPrint([FromBody] PrintModel print)
        {

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
                Document doc = new Document(PageSize.LETTER, 20, 20,70,40);

                PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(absolutePath, FileMode.Create));
                writer.PageEvent = new HeaderFooter();

                //Head table body
                Paragraph product_id = new Paragraph("Producto Id", boldFontRedHead);
                Paragraph product = new Paragraph("Producto", boldFontRedHead);
                Paragraph observation = new Paragraph("Observación", boldFontRedHead);

                doc.Open();

                //Table body
                PdfPTable body = new PdfPTable(3);
                body.HeaderRows = 1;
                body.AddCell(new PdfPCell(product_id) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0, PaddingBottom = 5, PaddingTop = 15 });
                body.AddCell(new PdfPCell(product) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0, PaddingBottom = 5, PaddingTop = 15 });
                body.AddCell(new PdfPCell(observation) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0, PaddingBottom = 5, PaddingTop = 15 });

                //Datos de las demas filas
                Paragraph product_id_value = new Paragraph("CERT", normalFont);
                Paragraph product_value = new Paragraph("Pellentesque consequat augue quis est interdum", normalFont);
                Paragraph observation_value = new Paragraph("Fusce risus dolor, accumsan et vestibulum eget", normalFont);

                //Mas filas
                body.AddCell(returnCell(product_id_value));
                body.AddCell(returnCell(product_value));
                body.AddCell(returnCell(observation_value));

                body.AddCell(returnCell(new Paragraph()));
                body.AddCell(returnCell(product_value));
                body.AddCell(returnCell(observation_value));

                body.AddCell(returnCell(new Paragraph()));
                body.AddCell(returnCell(product_value));
                body.AddCell(returnCell(observation_value));

                body.WidthPercentage = 100f;
                body.PaddingTop = 20;
                var colWidthPercentagesBody = new[] { 15f, 35f, 50f };
                body.SetWidths(colWidthPercentagesBody);

                doc.Add(body);
       
                doc.Close();
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

        public PdfPCell returnCell(Paragraph text)
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
        public bool IsPrinterOnline(string printerName)
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
