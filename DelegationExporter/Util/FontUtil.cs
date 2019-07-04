using DelegationExporter.Base;
using iText.IO.Font;
using iText.Kernel.Font;
using System;
using System.IO;

namespace DelegationExporter.Util
{
    public class FontUtil
    {
        public static PdfFont GetSystemFont()
        {
            string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "msjhbd.ttc,0");
            return PdfFontFactory.CreateFont(path, PdfEncodings.IDENTITY_H, true);
        }

        //微軟正黑體
        public static PdfFont GetMsjhbdFont()
        {         
            string path = Path.Combine(Config.FONT_FOLDER, "msjhbd.ttc,0");
            return PdfFontFactory.CreateFont(path, PdfEncodings.IDENTITY_H, true);
        }

        public static bool IsMsjhbdExist()
        {
            if(File.Exists(Path.Combine(Config.FONT_FOLDER, "msjhbd.ttc")))
            {
                return true;
            }
            return false;
        }
    }
}
