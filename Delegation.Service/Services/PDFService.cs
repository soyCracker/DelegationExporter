using iText.Forms.Fields;
using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Annot;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;

namespace Delegation.Service.Services
{
    public class PDFService
    {
        private string fontFolder;

        public PDFService(string fontFolder)
        {
            this.fontFolder = fontFolder;
        }

        //思源黑體 itext 7.1.5、7.1.14修改字體功能正常
        public PdfFont GetNotoSansFont()
        {
            string path = System.IO.Path.Combine(fontFolder, "NotoSansCJKtc-VF.ttf");
            PdfFont font = PdfFontFactory.CreateFont(path, PdfEncodings.IDENTITY_H, true);
            return font;
        }

        public bool IsNotoSansExist()
        {
            if (File.Exists(System.IO.Path.Combine(fontFolder, "NotoSansCJKtc-VF.ttf")))
            {
                return true;
            }
            return false;
        }

        //勾選框
        public void SetPdfCheckBoxSelected(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                // 設定字體、粗體
                Paragraph p = new Paragraph("v").SetFont(GetNotoSansFont()).SetBold();
                p.SetFixedPosition(rectangle.GetX(), rectangle.GetY() - 4, rectangle.GetWidth());
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }

        //寫入文字在輸入框中間
        public void SetPdfFeldValueCenter(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey, string content)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                // 設定字體、粗體
                Paragraph p = new Paragraph(content).SetFont(GetNotoSansFont()).SetBold();
                p.SetFixedPosition(rectangle.GetX() + (rectangle.GetWidth() / 3), rectangle.GetY() - 4, rectangle.GetWidth());
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }

        //寫入小文字在輸入框
        public void SetPdfFeldValueSmall(IDictionary<string, PdfFormField> fields, PdfDocument pdfDoc, string fieldsKey, string content)
        {
            foreach (PdfAnnotation annotation in fields[fieldsKey].GetWidgets())
            {
                PdfPage page = annotation.GetPage();
                Rectangle rectangle = annotation.GetRectangle().ToRectangle();
                Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, rectangle);
                // 設定字體、粗體
                Paragraph p = new Paragraph(content).SetFont(GetNotoSansFont()).SetFontSize(8).SetBold();
                p.SetFixedPosition(rectangle.GetX(), rectangle.GetY() - 2, rectangle.GetWidth());
                canvas.Add(p);
                canvas.Close();
                page.RemoveAnnotation(annotation);
            }
        }
    }
}