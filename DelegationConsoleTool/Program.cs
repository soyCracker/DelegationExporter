// See https://aka.ms/new-console-template for more information
using Delegation.Service.Services;
using DelegationConsoleTool.Base;
using DelegationConsoleTool.Kits;
using System.Text;

EnvirKit envirKit = new EnvirKit();
envirKit.InitEnvironment();
string delegationFormFolder = Constant.FILE_FOLDER + @"\" + Constant.DELEGATION_FORM_FOLDER;
string assignmentFolder = Constant.FILE_FOLDER + @"\" + Constant.ASSIGNMENT_FOLDER;

Console.WriteLine("Delegation Console Tool 功能:");
Console.WriteLine("委派紀錄填寫 - 1");
Console.WriteLine("輸出委派單 - 3");

Console.Write("輸入功能:");
string funKey = "";
funKey = Console.ReadLine();

if (funKey=="1")
{
    Console.WriteLine("委派紀錄填寫");
    RecordService recordService = new RecordService();
    recordService.Start(Constant.OUTPUT_FOLDER, envirKit.PrepareAndGetTempXlsx(delegationFormFolder), envirKit.GetAssignRecordXlsx(assignmentFolder));
}
