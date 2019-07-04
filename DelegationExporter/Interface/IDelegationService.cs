using DelegationExporter.Model;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace DelegationExporter.Interface
{
    public interface IDelegationService <T>
    {
        List<T> ReadDelegation(string filePath, string sheetName);

        /// <summary>
        /// 
        /// </summary>
        /// <param name="delegation"></param>
        /// <param name="destFolder"></param>
        /// <returns>person name, as pdf name</returns>
        void WriteDelegation(SongSanDelegationXlsxModel delegation, string destFolder);

    }
}
