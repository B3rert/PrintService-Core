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
                var normalFont = FontFactory.GetFont(FontFactory.HELVETICA, 10);
                var littleFont = FontFactory.GetFont(FontFactory.HELVETICA, 7);
                var littleFontBold = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 7);
                var normalFontRed = FontFactory.GetFont(FontFactory.HELVETICA, 10,color_red);
                var boldFontRed = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 12, color_red);
                var boldFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
               
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
                date_internal_control_value.Alignment = Element.ALIGN_CENTER;


                var phrase = new Phrase();
                phrase.Add(new Chunk("REASON(S) FOR CANCELLATION:", boldFont));
                phrase.Add(new Chunk(" See Statutoryreason(s) designated by Code No(s) 1 on the reverse side hereof", normalFont));
                
            
                doc.Open();


                PdfPTable table = new PdfPTable(4);

                

                // Esta es la primera fila
                table.AddCell(returnCell(client));
                table.AddCell(returnCell(client_name));
                table.AddCell(returnCell(new Paragraph()));
                table.AddCell(returnCell(new Paragraph()));

                // Segunda fila
                table.AddCell(returnCell(phone) ); 
                table.AddCell(returnCell(client_phone) );
                table.AddCell(returnCell(new Paragraph()));
                table.AddCell(returnCell(new Paragraph()));

                // Tercera fila
                table.AddCell(returnCell(email)); 
                table.AddCell(returnCell(client_email));
                table.AddCell(returnCell(no_internal_control));
                table.AddCell(returnCell(date_internal_control));
                // Cuarta fila
                table.AddCell(returnCell(address));
                table.AddCell(returnCell(client_address));
                table.AddCell(returnCell(no_internal_control_value));
                table.AddCell(returnCell(date_internal_control_value));

                table.DefaultCell.Border = Rectangle.NO_BORDER;
                table.WidthPercentage = 100f;
                var colWidthPercentages = new[] { 15f, 49f, 18f, 18f};
                table.SetWidths(colWidthPercentages);
                // Agregamos la tabla al documento
                //agrgamos tofo al docuemnto
                doc.Add(logo_empresa);
                doc.Add(logo_dev);
                doc.Add(title);
                doc.Add(paragraph_void);
                doc.Add(table);

                doc.Add(phrase);

                doc.Add(new Paragraph("\n\n\n\n\n"));
                doc.Add(new Paragraph(print.doc));
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
