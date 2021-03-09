
namespace DelegationExporter.Base
{
    public class Constant
    {
        //process mode
        public static readonly bool RELEASE_MODE = false;

        //program info
        public static readonly string VERSION = "1.0.6-dev-20210309";

        //url
        public static readonly string GIT_URL = "https://github.com/soyCracker/DelegationExporter/releases";
        public static readonly string GIT_RELEASE_API = "https://api.github.com/repos/soyCracker/DelegationExporter/releases";

        //file path
        public static readonly string FILE_FOLDER = "File";
        public static readonly string OUTPUT_FOLDER = "Export";
        public static readonly string CONFIG_FILE = "config.json";
        public static readonly string TARGET_XLSX = "S89.xlsx";
        public static readonly string BRO_SHEET = "tw-sonsan-bro";
        public static readonly string SIS_SHEET = "tw-sonsan-sis";
        public static readonly string FONT_FOLDER = "Font";
        public static readonly string TEMP_NAME = "TEMPX";
        public static readonly string PDF_FILE_NAME_EXTENSION = ".pdf";

        //S-89 file
        public static readonly string S89CH = "S-89-CH.pdf";
        public static readonly string S89J = "S-89-J.pdf";

        //others
        public static readonly string DESCRIP_STR = "傳道與生活聚會委派通知單-";
        public static readonly string DISTINT_JP_STR = "JP";
        public static readonly string DESCRIP_JP_STR = "日語-";
    }
}
