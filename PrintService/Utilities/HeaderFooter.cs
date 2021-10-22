using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace PrintService.Utilities
{
    public class HeaderFooter: PdfPageEventHelper
    {
        public override void OnEndPage(PdfWriter writer, Document document)
        {

            var currentDirectory = Directory.GetCurrentDirectory(); //Ruta donden se encuntra el programa

            PdfContentByte cb = writer.DirectContent; //PDF que está escribiendose


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
            Paragraph title_report = new Paragraph(Globales.title_report, boldFontRed);
            Paragraph name_emmited = new Paragraph(Globales.name_emited, boldFont);

            PdfPTable header = new PdfPTable(1);
            header.TotalWidth = document.Right - document.Left;

            header.AddCell(new PdfPCell(title_report) { HorizontalAlignment = Element.ALIGN_CENTER, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
            header.AddCell(new PdfPCell(new Paragraph(" ")) { BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
            header.AddCell(new PdfPCell(new Paragraph(" ")) { BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
            header.AddCell(new PdfPCell(name_emmited) { HorizontalAlignment = Element.ALIGN_LEFT, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });

            header.WriteSelectedRows(0, -1, 20, 790, writer.DirectContent);

            //Footer
            DateTime dateTime = DateTime.Now;
            var hora = dateTime.Hour;
            var suffix = (hora >= 12) ? " p.m." : " a.m.";
            string fecha_hora = dateTime.ToString("dd/MM/yyyy hh:mm:ss") + suffix;

            Phrase paginator = new Phrase();
            paginator.Add(new Chunk(fecha_hora, boldGray));
            paginator.Add(new Chunk($" Página {document.PageNumber}", normalFont9));

            Globales.text_info = Globales.text_info.Replace("\t","  ");

            Paragraph text_info = new Paragraph(Globales.text_info, boldBlue);

            PdfPTable footer = new PdfPTable(3);
            footer.TotalWidth = document.Right - document.Left;

            var colWidthPercentagesFooter = new[] { 40f, 50f, 10f };
            footer.SetWidths(colWidthPercentagesFooter);

            footer.AddCell(new PdfPCell(paginator) { HorizontalAlignment = Element.ALIGN_LEFT, VerticalAlignment = Element.ALIGN_BOTTOM, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
            footer.AddCell(new PdfPCell(text_info) { HorizontalAlignment = Element.ALIGN_CENTER, VerticalAlignment = Element.ALIGN_CENTER, BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });
            footer.AddCell(new PdfPCell(new Paragraph(" ")) { BorderWidthBottom = 0, BorderWidthLeft = 0, BorderWidthTop = 0, BorderWidthRight = 0 });

            footer.WriteSelectedRows(0, -1, 20, 40, writer.DirectContent);

        }
    }
}
