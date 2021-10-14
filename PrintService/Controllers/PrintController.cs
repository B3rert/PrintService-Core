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
            var currentDirectory = Directory.GetCurrentDirectory(); //Ruta donden se encuntra el programa
            var absolutePath = $"{currentDirectory}\\testprinxt.pdf";
            var pagesPdf = 1;

            try
            {
                //450 margen derecho
                Document doc = new Document();
                PdfWriter.GetInstance(doc, new FileStream(absolutePath, FileMode.Create));
                Image logo_empresa = Image.GetInstance("Assets\\empresa_logo.jfif");
                logo_empresa.ScalePercent(50f);
                logo_empresa.SetAbsolutePosition(450,725);
                Image logo_dev = Image.GetInstance("Assets\\demosoft.jfif");
                logo_dev.ScalePercent(10f);
                logo_dev.SetAbsolutePosition(525, 25);

                //image1.ScaleAbsoluteWidth(480);
                // image1.ScaleAbsoluteHeight(270);
                doc.Open();

                doc.Add(logo_empresa);
                doc.Add(logo_dev);
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
    }
}
