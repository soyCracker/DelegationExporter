using System.Collections.Generic;

namespace DelegationExporter.Interface
{
    public interface IDelegationService
    {
        //void DoWork();

        List<T> ReadDelegation<T>(string filePath);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="delegation"></param>
        /// <param name="destFolder"></param>
        /// <returns>person name, as pdf name</returns>
        void WriteDelegation<T>(List<T> delegationList, string destFolder);

        string BeforePrepareAndGetTempXlsx();
    }
}
