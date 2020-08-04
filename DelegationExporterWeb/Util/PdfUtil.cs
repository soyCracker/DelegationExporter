using DelegationExporterEntity.Entities;
using DelegationExporterWeb.Base;
using iText.Forms.Fields;
using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DelegationExporterWeb.Util
{
    public class PdfUtil
    {

        //微軟正黑體 itext 7.1.5修改字體功能正常
        /*public static PdfFont GetMsjhbdFont(DelegationExporterDBContext context)
        {
            NecessaryFile necessaryFile = context.NecessaryFile.FirstOrDefault(x => x.FileName == Constant.MSJHBD_TTC);
            return PdfFontFactory.CreateTtcFont(necessaryFile.Data, 0, PdfEncodings.IDENTITY_H, true, false);
        }

        public static void SetPdfCheckBoxSelected(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey, DelegationExporterDBContext context)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                Paragraph p = new Paragraph("v").SetFont(GetMsjhbdFont(context));
                p.SetFixedPosition(rectangle.GetX(), rectangle.GetY() - 4, rectangle.GetWidth()).SetFont(GetMsjhbdFont(context));
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }

        public static void SetPdfFeldValueCenter(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey, string content, DelegationExporterDBContext context)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                Paragraph p = new Paragraph(content).SetFont(GetMsjhbdFont(context));
                p.SetFixedPosition(rectangle.GetX() + (rectangle.GetWidth() / 3), rectangle.GetY() - 4, rectangle.GetWidth()).SetFont(GetMsjhbdFont(context));
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }

        public static void SetPdfFeldValueSmall(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey, string content, DelegationExporterDBContext context)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                Paragraph p = new Paragraph(content).SetFont(GetMsjhbdFont(context)).SetFontSize(8);
                p.SetFixedPosition(rectangle.GetX(), rectangle.GetY() - 2, rectangle.GetWidth()).SetFont(GetMsjhbdFont(context));
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }*/
    }
}
