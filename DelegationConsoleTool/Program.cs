using Delegation.Service.Except;
using Delegation.Service.Services;
using DelegationConsoleTool.Base;
using DelegationConsoleTool.Kits;
using System.Text;


EnvirKit envirKit = new EnvirKit();
envirKit.InitEnvironment();
string delegationFormFolder = Constant.FILE_FOLDER + @"\" + Constant.DELEGATION_FORM_FOLDER;
string assignmentFolder = Constant.FILE_FOLDER + @"\" + Constant.ASSIGNMENT_FOLDER;
string s89chFile = Constant.FILE_FOLDER + @"\" + Constant.S89CH;
string s89jpFile = Constant.FILE_FOLDER + @"\" + Constant.S89J;

while (true)
{
    FucDescOutput();
    string fucKey = Console.ReadLine();
    Fuchub(fucKey);
}

void FucDescOutput()
{
    Console.WriteLine("Release Mode: " + Constant.RELEASE_MODE);
    Console.WriteLine("Delegation Console Tool 功能:");
    Console.WriteLine("委派紀錄填寫 - 1");
    Console.WriteLine("輸出委派單 - 3");
    Console.Write("輸入功能:");
}

void Fuchub(string fucKey)
{
    if (fucKey == "1")
    {
        RecordWork();
    }
    else if (fucKey == "3")
    {
        ExportWork();
    }
    else if (fucKey=="99")
    {
        TestWork();
    }
}

void RecordWork()
{
    Console.WriteLine("委派紀錄填寫");
    RecordService recordService = new RecordService();
    recordService.Start(Constant.OUTPUT_FOLDER, envirKit.PrepareAndGetTempXlsx(delegationFormFolder), envirKit.GetAssignRecordXlsx(assignmentFolder));
}

void ExportWork()
{
    try
    {
        Console.WriteLine("輸出委派單");
        string xlsx = envirKit.PrepareAndGetTempXlsx(delegationFormFolder);
        if(string.IsNullOrEmpty(xlsx))
        {
            Console.WriteLine(Path.GetFullPath(delegationFormFolder) + " 沒有委派單可輸出\n");
            Console.WriteLine(Directory.GetCurrentDirectory() + " 沒有委派單可輸出\n");
        }
        else
        {
            ExportService exportService = new ExportService(new PDFService(Constant.FONT_FOLDER));
            exportService.Start(Constant.OUTPUT_FOLDER, xlsx, s89chFile,
                s89jpFile, Constant.DESC_STR, Constant.DESC_JP_STR, Constant.JP_FLAG_STR);
        }  
    }
    catch (IOException ex)
    {
        Console.WriteLine("可能excel或pdf檔案不存在、config內的檔名填錯，或是檔案正被使用中請關閉所有檔案再試一次\n");
        Console.WriteLine(ex + "\n");
    }
    catch (NoFontException ex)
    {
        Console.WriteLine("字型檔案不存在，請檢查Font資料夾或重新下載程式\n");
        Console.WriteLine("Update Web url:" + Constant.GIT_RELEASE_API + "\n");
        Console.WriteLine(ex + "\n");
    }
    catch (MultipleXlsException ex)
    {
        Console.WriteLine("存在多個xls或xlsx於資料夾中\n");
        Console.WriteLine(ex + "\n");
    }
    catch (Exception ex)
    {
        Console.WriteLine("未知錯誤，請複製以下資訊給開發者:" + ex + "\n");
    }
}

static void TestWork()
{
    Console.WriteLine("TestWork() !!!\n");
}





