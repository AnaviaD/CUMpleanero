using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CUMpleanero
{
    public class Constantes
    {

        #region Variables & Constantes
            public const string _xServidorBD = "SRVFTRDB1";//"SRVFTRDB1";
            public const string _xNombreBD = "X_FTR";
            public const string _xUsuarioBD = "sa";
            public const string _xPassWordBD = "mxsis86%1021&js";

            public const string _xServidorBD2 = "SRVFTRDB1";//"SRVFTRDB1";
            public const string _xNombreBD2 = "FTRAIL";
            public const string _xUsuarioBD2 = "sa";
            public const string _xPassWordBD2 = "mxsis86%1021&js";

            //public const string ServidorSFTP = "sftp.qs3.com.mx";
            //public const string UsuarioSFTP = "truiz";
            //public const string PassWordSFTP = "h@Zh7RyX";
            public const int PuertoSFTP = 21;

            public const string ServidorSFTP = "35.245.240.87";
            public const string UsuarioSFTP = "Chrysler";
            public const string PassWordSFTP = "9V$QpaEM";

            public const string RutaDescargaSFTP = "/outgoing/"; //"/tesout/";
            public const string RutaUploadSFTP = "/PROD/"; //"/testin/";
            public const string SourceDir = @"\\10.1.1.110\c$\Users\Public\Documents";
            public const string SourceDhlS = @"\\10.1.1.30\FTProot\sftr-dhlsupply";
        #endregion
    }
}