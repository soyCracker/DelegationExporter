using System.Collections.Generic;

namespace DelegationExporter.Interface
{
    public interface IDelegationService
    {
        void DoWork();

        List<T> ReadDelegation<T>(string filePath, string sheetName);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="delegation"></param>
        /// <param name="destFolder"></param>
        /// <returns>person name, as pdf name</returns>
        void WriteDelegation<T>(T delegation, string destFolder);

    }
}
