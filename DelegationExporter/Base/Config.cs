﻿
namespace DelegationExporter.Base
{
    public class Config
    {
        //process mode
        public static bool RELEASE_MODE = false;

        //program info
        public static string VERSION = "1.0.5-20200527";

        //url
        public static string GIT_URL = "https://github.com/soyCracker/DelegationExporter/releases";
        public static string GIT_RELEASE_API = "https://api.github.com/repos/soyCracker/DelegationExporter/releases";

        //file path
        public static string FILE_FOLDER = "File";
        public static string OUTPUT_FOLDER = "Export";
        public static string CONFIG_FILE = "config.json";
        public static string TARGET_XLSX = "S89.xlsx";
        public static string BRO_SHEET = "tw-sonsan-bro";
        public static string SIS_SHEET = "tw-sonsan-sis";
        public static string FONT_FOLDER = "Font";
        public static string TEMP_NAME = "TEMPX";

        //S-89 file
        public static string S89CH = "S-89-CH.pdf";
        public static string S89J = "S-89-J.pdf";
    }
}
