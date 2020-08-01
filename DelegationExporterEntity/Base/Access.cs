using System;
using System.Collections.Generic;
using System.Text;

namespace DelegationExporterEntity.Base
{
    public class Access
    {
        public static readonly string DB_CONN = "Server=tcp:yuhui.database.windows.net,1433;Initial Catalog = DelegationExporterDB; Persist Security Info=False;User ID = yuhui; Password=PassYu@2020;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout = 30;";

        //public static readonly string DB_CONN = "data source=172.104.93.64;initial catalog=DelegationExporterDB;persist security info=True;user id=sa;password=Gaussian1998@gmail.com;Connection Timeout=120;MultipleActiveResultSets=True;App=EntityFramework";
    }
}
