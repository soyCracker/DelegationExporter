using DelegationExporter.Base;
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
using System.IO;

namespace DelegationExporter.Util
{
    public class PdfUtil
    {
        public static PdfFont GetSystemFont()
        {
            string path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "msjhbd.ttc,0");
            return PdfFontFactory.CreateFont(path, PdfEncodings.IDENTITY_H, true);
        }

        //微軟正黑體 itext 7.1.5修改字體功能正常
        public static PdfFont GetMsjhbdFont()
        {         
            string path = System.IO.Path.Combine(Constant.FONT_FOLDER, "msjhbd.ttc,0");
            return PdfFontFactory.CreateFont(path, PdfEncodings.IDENTITY_H, true);
        }

        public static bool IsMsjhbdExist()
        {
            if(File.Exists(System.IO.Path.Combine(Constant.FONT_FOLDER, "msjhbd.ttc")))
            {
                return true;
            }
            return false;
        }

        public static void SetPdfCheckBoxSelected(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                Paragraph p = new Paragraph("v").SetFont(GetMsjhbdFont());
                p.SetFixedPosition(rectangle.GetX(), rectangle.GetY() - 4, rectangle.GetWidth()).SetFont(GetMsjhbdFont());
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }

        public static void SetPdfFeldValueCenter(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey, string content)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                Paragraph p = new Paragraph(content).SetFont(GetMsjhbdFont());
                p.SetFixedPosition(rectangle.GetX() + (rectangle.GetWidth() / 3), rectangle.GetY() - 4, rectangle.GetWidth()).SetFont(GetMsjhbdFont());
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }

        public static void SetPdfFeldValueSmall(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey, string content)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                Paragraph p = new Paragraph(content).SetFont(GetMsjhbdFont()).SetFontSize(8);
                p.SetFixedPosition(rectangle.GetX(), rectangle.GetY() - 2, rectangle.GetWidth()).SetFont(GetMsjhbdFont());
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }

        public static string FileNameExistAddR(string filePath)
        {
            string result = filePath;
            while (true)
            {
                if (!File.Exists(result))
                {
                    return result;
                }
                result = result.Replace(".pdf","R") + ".pdf";
            }
        }
    }
}
