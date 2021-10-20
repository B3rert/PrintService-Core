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
        /*
        protected void Page_Load(object sender, EventArgs e)
        {
        
            using (MemoryStream ms = new MemoryStream())
            {
  
                 iTextSharp.text.Document doc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 36, 36, 54, 54);
                iTextSharp.text.pdf.PdfWriter writer = iTextSharp.text.pdf.PdfWriter.GetInstance(doc, ms);
                writer.PageEvent = new HeaderFooter();
                doc.Open();
  
                 // make your document content..
        
                doc.Close();
                writer.Close();

                // output
                Response.ContentType = "application/pdf;";
                Response.AddHeader("Content-Disposition", "attachment; filename=clientfilename.pdf");
                byte[] pdf = ms.ToArray();
                Response.OutputStream.Write(pdf, 0, pdf.Length);
                
            }
 
        }

        class HeaderFooter : PdfPageEventHelper
        {
            public override void OnEndPage(PdfWriter writer, Document document)
            {

                // Make your table header using PdfPTable and name that tblHeader
    
                tblHeader.WriteSelectedRows(0, -1, page.Left + document.LeftMargin, page.Top, writer.DirectContent);
    
                // Make your table footer using PdfPTable and name that tblFooter
    
                tblFooter.WriteSelectedRows(0, -1, page.Left + document.LeftMargin, writer.PageSize.GetBottom(document.BottomMargin), writer.DirectContent);
            }
        }

        */
        //Genera un PDF con el texto recibido y lo imprime en la impreosra especificada
        [HttpPost("generate")]
        public IActionResult getPrint([FromBody] PrintModel print)
        {

            bool result =  IsPrinterOnline(print.printer);
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
                var normalFontRed = FontFactory.GetFont(FontFactory.HELVETICA, 10,color_red);
                var boldFontRed = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, color_red);
                var boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
                var boldFontRedHead = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10,color_red);
                var boldGray = FontFactory.GetFont(FontFactory.HELVETICA,8,color_gray);
                var boldBlue = FontFactory.GetFont(FontFactory.HELVETICA,7,color_blue);


                //Crear el documento




                Document doc = new Document();
                PdfWriter.GetInstance(doc, new FileStream(absolutePath, FileMode.Create));

                //Parrafo vacio
                Paragraph paragraph_void = new Paragraph(" ");


                //Logos
                Image logo_empresa = Image.GetInstance($"{currentDirectory}\\assets\\empresa_logo.jfif");
                logo_empresa.ScalePercent(50f);
                logo_empresa.SetAbsolutePosition(450, 725);
                Image logo_dev = Image.GetInstance($"{currentDirectory}\\assets\\demosoft.jfif");
                logo_dev.ScalePercent(10f);
                logo_dev.SetAbsolutePosition(525, 25);

                //Titulo
                Paragraph title = new Paragraph("Receta", boldFontRed);
                title.Alignment = 1;

                //Encabezados
                Paragraph client = new Paragraph("CLIENTE", boldFont);
                Paragraph phone = new Paragraph("TELEFONO", boldFont);
                Paragraph email = new Paragraph("EMAIL", boldFont);
                Paragraph address = new Paragraph("DIRECCION", boldFont);

                //Datos clientes (Encabezados)
                Paragraph client_name = new Paragraph("DESARROLLO MODERNO DE SOFTWARE, S.A.", normalFont);
                Paragraph client_phone = new Paragraph("56918326", normalFont);
                Paragraph client_email = new Paragraph("gerencia@demosoftonline.com", normalFont);
                Paragraph client_address = new Paragraph("CIUDAD", normalFont);

                //Control interno
                Paragraph no_internal_control = new Paragraph("NO. Control Interno", littleFontBold);
                no_internal_control.Alignment = Element.ALIGN_CENTER;
                Paragraph date_internal_control = new Paragraph("Fecha de Control Interno", littleFontBold);
                date_internal_control.Alignment = Element.ALIGN_CENTER;
                Paragraph no_internal_control_value = new Paragraph("19", littleFontBold);
                no_internal_control_value.Alignment = Element.ALIGN_CENTER;
                Paragraph date_internal_control_value = new Paragraph("12/12/2021", littleFontBold);

                //Head table body
                Paragraph product_id = new Paragraph("Producto Id", boldFontRedHead);
                Paragraph product = new Paragraph("Producto", boldFontRedHead);
                Paragraph observation = new Paragraph("Observación", boldFontRedHead);

                //Foter
                Phrase coment = new Phrase();
                coment.Add(new Chunk("Comentario: ", boldFont));
                coment.Add(new Chunk("See Statutoryreason(s) designated by Code No(s) 1 on the reverse side hereof", normalFont9));

                Paragraph name_tended = new Paragraph("Atendió: Dr. Cabrera.",boldFont);
                Paragraph name_report = new Paragraph("file:///C:/Users/dsdev/Downloads/TD_Receta.pdf",boldGray);
                Phrase paginator = new Phrase();
                paginator.Add(new Chunk("19/10/2021 11:50:04 a.m.",boldGray));
                paginator.Add(new Chunk("Página 1 de 1",normalFont9));
                Paragraph text_info = new Paragraph("PBX: 2259-3232 / 6a. Ave \"A\" 13-25 Zona 9, Guatemala/ info@imcguate.com\ndrcabreramancio@imcguate.com    5552-417    /IMCCabreraMancio/\nwww.imcguate.com",boldBlue);






                doc.Open();


                PdfPTable header = new PdfPTable(4);

                // Esta es la primera fila
                header.AddCell(returnCell(client));
                header.AddCell(returnCell(client_name));
                header.AddCell(returnCell(new Paragraph()));
                header.AddCell(returnCell(new Paragraph()));

                // Segunda fila
                header.AddCell(returnCell(phone) );
                header.AddCell(returnCell(client_phone) );
                header.AddCell(returnCell(new Paragraph()));
                header.AddCell(returnCell(new Paragraph()));

                // Tercera fila
                header.AddCell(returnCell(email));
                header.AddCell(returnCell(client_email));
                header.AddCell(new PdfPCell(no_internal_control) { HorizontalAlignment = Element.ALIGN_CENTER, VerticalAlignment = Element.ALIGN_BOTTOM, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                header.AddCell(new PdfPCell(date_internal_control) { HorizontalAlignment = Element.ALIGN_CENTER, VerticalAlignment = Element.ALIGN_BOTTOM, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                // Cuarta fila
                header.AddCell(returnCell(address));
                header.AddCell(returnCell(client_address));
                header.AddCell(new PdfPCell(no_internal_control_value) { HorizontalAlignment = Element.ALIGN_CENTER, VerticalAlignment = Element.ALIGN_TOP, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                header.AddCell(new PdfPCell(date_internal_control_value) { HorizontalAlignment = Element.ALIGN_CENTER, VerticalAlignment = Element.ALIGN_TOP, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });

                //header.DefaultCell.Border = Rectangle.NO_BORDER;
                header.WidthPercentage = 100f;
                var colWidthPercentages = new[] { 15f, 49f, 18f, 18f};
                header.SetWidths(colWidthPercentages);

                //Table body
                PdfPTable body = new PdfPTable(3);
                body.AddCell(new PdfPCell(product_id) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0,PaddingBottom = 5, PaddingTop = 15});
                body.AddCell(new PdfPCell(product) {HorizontalAlignment = Element.ALIGN_LEFT,  BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0, PaddingBottom = 5, PaddingTop = 15 });
                body.AddCell(new PdfPCell(observation) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 1, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0, PaddingBottom = 5, PaddingTop = 15 });

                //Datos de las demas filas
                Paragraph product_id_value = new Paragraph("CERT",normalFont);
                Paragraph product_value = new Paragraph("Pellentesque consequat augue quis est interdum", normalFont);
                Paragraph observation_value = new Paragraph("Fusce risus dolor, accumsan et vestibulum eget", normalFont);

                //Mas filas
                body.AddCell(returnCell(product_id_value));
                body.AddCell(returnCell(product_value));
                body.AddCell(returnCell(observation_value));

                body.AddCell(returnCell(product_id_value));
                body.AddCell(returnCell(product_value));
                body.AddCell(returnCell(observation_value));

                body.WidthPercentage = 100f;
                body.PaddingTop = 20;
                var colWidthPercentagesBody = new[] { 15f, 35f, 50f};
                body.SetWidths(colWidthPercentagesBody);


                //Footer
                PdfPTable footer = new PdfPTable(3);
                footer.WidthPercentage = 100f;
                
                var colWidthPercentagesFooter = new[] {  45f, 50f, 5f };
                footer.SetWidths(colWidthPercentagesFooter);

                //content table footer 
                footer.AddCell(returnCell(name_tended));
                footer.AddCell(new PdfPCell(name_report) { HorizontalAlignment = Element.ALIGN_CENTER, VerticalAlignment = Element.ALIGN_CENTER, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                footer.AddCell(returnCell(new Paragraph()));

                //
                footer.AddCell(new PdfPCell(paginator) { HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_BOTTOM, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                //footer.AddCell(returnCell(paginator));
                footer.AddCell(new PdfPCell(text_info) { HorizontalAlignment = Element.ALIGN_CENTER, VerticalAlignment = Element.ALIGN_CENTER, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
                footer.AddCell(returnCell(new Paragraph()));


                //Agregar contenido al documento

                doc.Add(logo_empresa);
                doc.Add(logo_dev);
                doc.Add(title);
                doc.Add(paragraph_void);
                doc.Add(header);
                doc.Add(body);
                doc.Add(new Paragraph("\n\n\n\n\n"));
                doc.Add(coment);
                doc.Add(footer);
              //  doc.Add(new Paragraph(print.doc));
                //doc.Add(new Paragraph("The standard chunk of Lorem Ipsum used since the 1500s is reproduced below for those interested. Sections 1.10.32 and 1.10.33 from de Finibus Bonorum et Malorum by Cicero are also reproduced in their exact original form, accompanied by English versions from the 1914 translation by H. Rackham."));

                doc.Close();
            }
            catch (Exception e)
            {
                return BadRequest(e.Message);
            }

           
            /*

            //Obtiene el numero de páginas
            using (StreamReader sr = new StreamReader(System.IO.File.OpenRead(absolutePath)))
            {
                Regex regex = new Regex(@"/Type\s*//*Page[^s]");
                MatchCollection matches = regex.Matches(sr.ReadToEnd());
                pagesPdf = matches.Count;
            }

            */
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
