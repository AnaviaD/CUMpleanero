#region Using
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Data.Common;
    using System.Data.SqlClient;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Net;
    using System.Resources;
    using System.Runtime.CompilerServices;
    using System.Threading;
    using System.Windows.Forms;
    using System.Text;
    using System.Text.RegularExpressions;
    using WinSCP;
using OfficeOpenXml;
#endregion

namespace FTPALTOMOVUP
{
    public partial class FTPAltoMovup : Form
    {
        #region Variables&Constantes
            private SqlConnection _xConnString = new SqlConnection();
            private SqlConnection _xConnString2 = new SqlConnection();
            private DirectoryInfo di;
            private List<string> xRenglon = new List<string>();
            private List<ArchivosDescarga> xTracksIDS = new List<ArchivosDescarga>();
            //private WebClient _xWebClient;
            //private ArchivosDescarga _ArchivosDescarga;
            private string _xFolderPath;
            //private string _xNombreXML;
            //private string _xHTMLUser;
            //private string _xHTMLPass;
            private int _xManual;
            private int _xError;
        #endregion

        #region Constructor
            public FTPAltoMovup()
            {
                InitializeComponent();
                this.CargarConexion(Constantes._xServidorBD, Constantes._xNombreBD, Constantes._xUsuarioBD, Constantes._xPassWordBD);
                this.CargarConexion2(Constantes._xServidorBD2, Constantes._xNombreBD2, Constantes._xUsuarioBD2, Constantes._xPassWordBD2);
            }
        #endregion

        #region Eventos
            private void btn_actuaizar_Click(object sender, EventArgs e)
            {
                System.Windows.Forms.Cursor cursor = this.Cursor;
                this.Cursor = Cursors.WaitCursor;
                try
                {
                    this.btn_actuaizar.Enabled = false;
                    this.pgb_Progreso.Value = 0;
                    this.tmr_manual.Start();
                    this._xManual = 1;
                    this.pgb_Progreso.Value = 0;
                    this.Cursor = Cursors.Default;
                }
                catch
                {
                    this.ltb_log.Items.Add(string.Format("{0} - Error al actualizar la información, Intentelo Nuevamente...", DateTime.Now));
                    this.ltb_log.Items.Add("");
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                    this.Cursor = Cursors.Default;
                }
            }
        #endregion

        #region Metodos Privados
            private void SFTP_Load(object sender, EventArgs e)
            {
                this.IniciarTimer();
            }
            private void CargarConexion(string _xServerName, string _xNombreBD, string _xUser, string _xPassword)
            {
                this._xConnString.ConnectionString = string.Format("data source = {0}; initial catalog = {1}; User Id={2}; Password = {3};", new object[] { _xServerName, _xNombreBD, _xUser, _xPassword });
            }
            private void CargarConexion2(string _xServerName, string _xNombreBD, string _xUser, string _xPassword)
            {
                this._xConnString2.ConnectionString = string.Format("data source = {0}; initial catalog = {1}; User Id={2}; Password = {3};", new object[] { _xServerName, _xNombreBD, _xUser, _xPassword });
            }
            private void IncreaseProgressBarPosiciones(object sender, EventArgs e)
            {
                this.GetClientFiles_Message();
                if (this.pgb_Progreso.Value == this.pgb_Progreso.Maximum)
                {
                    this.Proceso();
                    this.pgb_Progreso.Value = 0;
                    Thread.Sleep(4000);
                }
                else if (this.pgb_Progreso.Value == 90)
                {
                    //this.CargarLogIn();
                }
                else if (this.pgb_Progreso.Value == 95)
                {
                    //this.CargarTracks();
                }
                this.pgb_Progreso.Increment(1);
            }
            private void IniciarTimer()
            {
                this.tmr_Actualiza.Interval = 3600;
                this.tmr_Actualiza.Tick += new EventHandler(this.IncreaseProgressBarPosiciones);
            }
            private void Proceso()
            {
                System.Windows.Forms.Cursor cursor = this.Cursor;
                this.Cursor = Cursors.WaitCursor;
                this.ltb_log.Items.Add(string.Format("{0} - Sesión iniciada correctamente..", DateTime.Now));
                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                this._xError = 0;
                this.DirectorioDescarga();
                this.pgb_Progreso.Value = 0;
                this.GetFiles();
                this.pgb_Progreso.Value = 0;
                this.Cursor = cursor;
            }


        private void InsertaMercanciaClientes(string ReferenciaDelServicio, string RSdelRemitente, string RFCdelRemitente, string Supplier, string Calle, string Municipio, string Estado, string Pais, string CP, string RSdelDestinatario, string RFCDestinatario, string Calle2, string Municipio2, string Estado2, string Pais2, string CP2, string PesoNeto, string NumeroTotalMercancias, string ClaveDelBienTransportado, string DescripcionDelBienTransportado, string ClaveUnidadDeMedida, string MaterialPeligroso, string ValorDeLaMercancia, string TipoDeMoneda)
        {
            try
            {
                SqlCommand sqlCommand = new SqlCommand("SP_InsertaMercanciaClientes", this._xConnString)
                {
                    CommandType = CommandType.StoredProcedure
                };


                sqlCommand.Parameters.AddWithValue("@ReferenciaDelServicio", ReferenciaDelServicio);
                sqlCommand.Parameters.AddWithValue("@RSdelRemitente", RSdelRemitente);
                sqlCommand.Parameters.AddWithValue("@RFCdelRemitente", RFCdelRemitente);
                sqlCommand.Parameters.AddWithValue("@Supplier", Supplier);
                sqlCommand.Parameters.AddWithValue("@Calle", Calle);
                sqlCommand.Parameters.AddWithValue("@Municipio", Municipio);
                sqlCommand.Parameters.AddWithValue("@Estado", Estado);
                sqlCommand.Parameters.AddWithValue("@Pais", Pais);
                sqlCommand.Parameters.AddWithValue("@CP", CP);
                sqlCommand.Parameters.AddWithValue("@RSdelDestinatario", RSdelDestinatario);
                sqlCommand.Parameters.AddWithValue("@RFCDestinatario", RFCDestinatario);
                sqlCommand.Parameters.AddWithValue("@Calle2", Calle2);
                sqlCommand.Parameters.AddWithValue("@Municipio2", Municipio2);
                sqlCommand.Parameters.AddWithValue("@Estado2", Estado2);
                sqlCommand.Parameters.AddWithValue("@Pais2", Pais2);
                sqlCommand.Parameters.AddWithValue("@CP2", CP2);
                sqlCommand.Parameters.AddWithValue("@PesoNeto", PesoNeto);
                sqlCommand.Parameters.AddWithValue("@NumeroTotalMercancias", NumeroTotalMercancias);
                sqlCommand.Parameters.AddWithValue("@ClaveDelBienTransportado", ClaveDelBienTransportado);
                sqlCommand.Parameters.AddWithValue("@DescripcionDelBienTransportado", DescripcionDelBienTransportado);
                sqlCommand.Parameters.AddWithValue("@ClaveUnidadDeMedida", ClaveUnidadDeMedida);
                sqlCommand.Parameters.AddWithValue("@MaterialPeligroso", MaterialPeligroso);
                sqlCommand.Parameters.AddWithValue("@ValorDeLaMercancia", ValorDeLaMercancia);
                sqlCommand.Parameters.AddWithValue("@TipoDeMoneda", TipoDeMoneda);


                this._xConnString.Open();
                sqlCommand.ExecuteNonQuery();
                this._xConnString.Close();
            }
            catch
            {
                this.ltb_log.Items.Add(string.Format("{0} - Problemas al insertar los archivos.", DateTime.Now));
                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
        }



        private void GetClientFiles_Message()
        {

            string _fName = "";
            string str = "";
            string strFullNme = "";



            string ReferenciaDelServicio = "";
            string RSdelRemitente = "";
            string RFCdelRemitente = "";
            string Supplier = "";
            string Calle = "";
            string Municipio = "";
            string Estado = "";
            string Pais = "";
            string CP = "";
            string RSdelDestinatario = "";
            string RFCDestinatario = "";
            string Calle2 = "";
            string Municipio2 = "";
            string Estado2 = "";
            string Pais2 = "";
            string CP2 = "";
            string PesoNeto = "";
            string NumeroTotalMercancias = "";
            string ClaveDelBienTransportado = "";
            string DescripcionDelBienTransportado = "";
            string ClaveUnidadDeMedida = "";
            string MaterialPeligroso = "";
            string ValorDeLaMercancia = "";
            string TipoDeMoneda = "";

            int tb_EmbarqueDHL = 1;
            int tb_Orden = 2;
            int tb_IDOrigen = 3;
            int tb_RFCRemitente2 = 4;
            int tb_Calle = 5;
            int tb_Municipio = 6;
            int tb_Estado = 7;
            int tb_Pais = 8;
            int tb_CodigoPostal = 9;
            int tb_IDDestino = 10;
            int tb_RFCDestinatario2 = 11;
            int tb_Calle2 = 12;
            int tb_Municipio3 = 13;
            int tb_Estado4 = 14;
            int tb_Pais5 = 15;
            int tb_CodigoPostal6 = 16;
            int tb_PesoNetoTotal = 17;
            int tb_NumTotalMercancias = 18;
            int tb_BienesTransp = 19;
            int tb_Descripción = 20;
            int tb_ClaveUnidad = 21;
            int tb_CveMaterialPeligroso = 22;
            int tb_Embalaje = 23;
            int tb_DescripEmbalaje = 24;
            int tb_ValorMercancia = 25;
            int tb_FraccionArancelaria = 26;
            int tb_UUIDComercioExt = 27;
            int tb_TotalKMRuta = 28;



Session session = new Session();
            try
            {
                // Setup session options
                SessionOptions sessionOptions = new SessionOptions
                {
                    Protocol = Protocol.Ftp,
                    FtpSecure = FtpSecure.Explicit,
                    FtpMode = FtpMode.Passive,
                    PortNumber = Constantes.PuertoSFTP,
                    HostName = Constantes.ServidorSFTP, //this.txtHostName.Text,
                    UserName = Constantes.UsuarioSFTP, //this.txtUserName.Text,
                    Password = Constantes.PassWordSFTP, //this.txtPwd.Text,
                    GiveUpSecurityAndAcceptAnySshHostKey = false,
                    GiveUpSecurityAndAcceptAnyTlsHostCertificate = true
                };
                session.Open(sessionOptions);

                bool exists = System.IO.Directory.Exists(@"\\10.1.1.30\FTProot\Done");

                if (!exists)
                    System.IO.Directory.CreateDirectory(@"\\10.1.1.30\FTProot\Done");

                var directories = Directory.GetDirectories(@"\\10.1.1.30\FTProot");
                for (int i = 1; i < (int)directories.Length; i++)
                {
                    FileInfo[] files = (new DirectoryInfo(directories[i])).GetFiles();
                    if ((directories[i].Contains("sftr-") ? true : directories[i].Contains("sftr-")))
                    {
                        for (int j = 0; j < (int)files.Length; j++)
                        { 
                            FileInfo fileInfo = files[j];
                            if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                            {
                                String filenameToUpload = fileInfo.Name;
                                _fName = fileInfo.Name;
                                strFullNme = fileInfo.FullName;
                                if (session.Opened)
                                {
                                    DataTable dt = new DataTable();
                                    FileStream fileStream = null;

                                    // limpieza ante todo
                                    dt.Clear();

                                    FileInfo datos = new FileInfo(fileInfo.DirectoryName + "\\" + fileInfo.Name);
                                    ExcelPackage LecExcel = new ExcelPackage(datos);
                                    try
                                    {
                                        fileStream = datos.Open(FileMode.Open, FileAccess.ReadWrite);
                                    }
                                    catch (NullReferenceException)
                                    {
                                        MessageBox.Show("El archivo se encuentra abierto por otro proceso.");
                                        return;
                                    }
                                    catch (IOException)
                                    {
                                        MessageBox.Show("El archivo se encuentra abierto por otro proceso.");
                                        return;
                                    }
                                    LecExcel.Load(fileStream);

                                    ExcelWorksheet worksheet = LecExcel.Workbook.Worksheets[1];
                                    fileStream.Close();

                                    if (worksheet.Dimension == null)
                                    {
                                        return;
                                    }

                                    //create a list to hold the column names
                                    List<string> columnNames = new List<string>();

                                    //needed to keep track of empty column headers
                                    int rows = worksheet.Dimension.Rows + 1;

                                    for (int e = 2; e < rows; e++)
                                    {
                                        if (worksheet.Cells[e, tb_EmbarqueDHL].Value?.ToString().Trim() != null)
                                        {

                                            ReferenciaDelServicio = worksheet.Cells[e, tb_EmbarqueDHL].Value?.ToString().Trim();
                                            RSdelRemitente = worksheet.Cells[e, tb_IDOrigen].Value?.ToString().Trim();
                                            RFCdelRemitente = worksheet.Cells[e, tb_RFCRemitente2].Value?.ToString().Trim();
                                            Supplier = worksheet.Cells[e, tb_RFCRemitente2].Value?.ToString().Trim();
                                            Calle = worksheet.Cells[e, tb_Calle].Value?.ToString().Trim();
                                            Municipio = worksheet.Cells[e, tb_Municipio].Value?.ToString().Trim();
                                            Estado = worksheet.Cells[e, tb_Estado].Value?.ToString().Trim();
                                            Pais = worksheet.Cells[e, tb_Pais].Value?.ToString().Trim();
                                            CP = worksheet.Cells[e, tb_CodigoPostal].Value?.ToString().Trim();
                                            RSdelDestinatario = worksheet.Cells[e, tb_IDDestino].Value?.ToString().Trim();
                                            RFCDestinatario = worksheet.Cells[e, tb_RFCDestinatario2].Value?.ToString().Trim();
                                            Calle2 = worksheet.Cells[e, tb_Calle2].Value?.ToString().Trim();
                                            Municipio2 = worksheet.Cells[e, tb_Municipio3].Value?.ToString().Trim();
                                            Estado2 = worksheet.Cells[e, tb_Estado4].Value?.ToString().Trim();
                                            Pais2 = worksheet.Cells[e, tb_Pais5].Value?.ToString().Trim();
                                            CP2 = worksheet.Cells[e, tb_CodigoPostal6].Value?.ToString().Trim();
                                            PesoNeto = worksheet.Cells[e, tb_PesoNetoTotal].Value?.ToString().Trim();
                                            NumeroTotalMercancias = worksheet.Cells[e, tb_NumTotalMercancias].Value?.ToString().Trim();
                                            ClaveDelBienTransportado = worksheet.Cells[e, tb_BienesTransp].Value?.ToString().Trim();
                                            ClaveUnidadDeMedida = worksheet.Cells[e, tb_ClaveUnidad].Value?.ToString().Trim();
                                            DescripcionDelBienTransportado = worksheet.Cells[e, tb_Descripción].Value?.ToString().Trim();
                                            MaterialPeligroso = worksheet.Cells[e, tb_CveMaterialPeligroso].Value?.ToString().Trim();
                                            ValorDeLaMercancia = worksheet.Cells[e, tb_ValorMercancia].Value?.ToString().Trim();
                                            TipoDeMoneda = worksheet.Cells[e, tb_FraccionArancelaria].Value?.ToString().Trim();

                                            InsertaMercanciaClientes(ReferenciaDelServicio, RSdelRemitente, RFCdelRemitente, Supplier, Calle, Municipio, Estado, Pais, CP, RSdelDestinatario, RFCDestinatario, Calle2, Municipio2, Estado2, Pais2, CP2, PesoNeto, NumeroTotalMercancias, ClaveDelBienTransportado, DescripcionDelBienTransportado, ClaveUnidadDeMedida, MaterialPeligroso, ValorDeLaMercancia, TipoDeMoneda);

                                        }
                                    }

                                }
                                else
                                {
                                    MessageBox.Show("Conexion error....", "FTP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                    
                        }


                        bool direxists = System.IO.Directory.Exists(@"\\10.1.1.30\FTPDon\Done\" + new DirectoryInfo(directories[i]).Name);

                        if (!direxists)
                            System.IO.Directory.CreateDirectory(@"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(directories[i]).Name);

                        string fileCutName = "";
                        string sourcePath = directories[i];
                        string targetPath = @"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(directories[i]).Name;

                        // Use Path class to manipulate file and directory paths.
                        string destFile = System.IO.Path.Combine(targetPath, fileCutName);


                        if (System.IO.Directory.Exists(sourcePath))
                        {
                            string[] filesIter = System.IO.Directory.GetFiles(sourcePath);

                            // Copy the files Iterand overwrite destination files Iterif they already exist.
                            foreach (string s in filesIter)
                            {
                                fileCutName = System.IO.Path.GetFileName(s);
                                if ((fileCutName.Contains(".xlsx") ? true : fileCutName.Contains(".xlsx")))
                                {
                                    // Use static Path methods to extract only the file name from the path.
                                    destFile = System.IO.Path.Combine(targetPath, fileCutName);
                                    System.IO.File.Copy(s, destFile, true);
                                    try
                                    {
                                        System.IO.File.Delete(s);
                                    }
                                    catch (System.IO.IOException e)
                                    {
                                        Console.WriteLine(e.Message);
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Source path does not exist!");
                    }

                }

                session.Close();
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                if (!File.Exists(str))
                {
                    File.Move(strFullNme, str);
                }
                else if (File.Exists(str))
                {
                    File.Delete(strFullNme);
                }
                session.Close();
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        /// <summary>
        /// Procesa el archivo de excel y lo carga en el Datatable
        /// </summary>
        /// <param name="dt"></param>
        private void ProcesoExcel(ref DataTable dt)
        {
            FileStream fileStream = null;

            // limpieza ante todo
            dt.Clear();

            FileInfo datos = new FileInfo(@"\\10.1.1.30\FTProot\sftr-dhlsupply");
            ExcelPackage LecExcel = new ExcelPackage(datos);
            try
            {
                fileStream = datos.Open(FileMode.Open, FileAccess.ReadWrite);
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("El archivo se encuentra abierto por otro proceso.");
                return;
            }
            catch (IOException)
            {
                MessageBox.Show("El archivo se encuentra abierto por otro proceso.");
                return;
            }
            LecExcel.Load(fileStream);

            ExcelWorksheet worksheet = LecExcel.Workbook.Worksheets[1];
            fileStream.Close();

            if (worksheet.Dimension == null)
            {
                return;
            }

            //create a list to hold the column names
            List<string> columnNames = new List<string>();

            //needed to keep track of empty column headers
            int currentColumn = 1;

            //loop all columns in the sheet and add them to the datatable
            foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
            {
                string columnName = cell.Text.Trim();

                //check if the previous header was empty and add it if it was
                if (cell.Start.Column != currentColumn)
                {
                    columnNames.Add("Header_" + currentColumn);
                    dt.Columns.Add("Header_" + currentColumn);
                    currentColumn++;
                }

                //add the column name to the list to count the duplicates
                columnNames.Add(columnName);

                //count the duplicate column names and make them unique to avoid the exception
                //A column named 'Name' already belongs to this DataTable
                int occurrences = columnNames.Count(x => x.Equals(columnName));
                if (occurrences > 1)
                {
                    columnName = columnName + "_" + occurrences;
                }

                //add the column to the datatable
                dt.Columns.Add(columnName);

                currentColumn++;
            }

            for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                DataRow newRow = dt.NewRow();

                //loop all cells in the row
                foreach (var cell in row)
                {
                    newRow[cell.Start.Column - 1] = cell.Text;
                }

                dt.Rows.Add(newRow);
            }


            var lista = new List<string>();
            string msg = string.Format("Mensaje de prueba Marcaje-Presión.\r\n\n ");

            lista.Add("desaftr04@ftr.com.mx");
            lista.Add("mmartinez@ftr.com.mx");
            //Notificar(lista, msg);
        }

        private void tmr_manual_Tick(object sender, EventArgs e)
            {
                try
                {
                    if (this._xManual == 1)
                    {
                        this._xManual = 2;
                    }
                    else if (this._xManual == 2)
                    {
                        this._xManual = 3;
                    }
                    else if (this._xManual == 3)
                    {
                        this.Proceso();
                        this._xManual = 0;
                        this.tmr_manual.Stop();
                        this.UploadFiles();
                        Thread.Sleep(4000);
                        this.btn_actuaizar.Enabled = true;
                    }
                }
                catch
                {
                    this.ltb_log.Items.Add(string.Format("{0} - Error al actualizar la información, Intentelo Nuevamente...", DateTime.Now));
                    this.ltb_log.Items.Add("");
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                    this.btn_actuaizar.Enabled = true;
                }
            }
            private void DirectorioDescarga()
            {
                string str = DateTime.Now.ToString("MMMM");
                DateTime now = DateTime.Now;
                string str1 = string.Format("{0}{1}", str, now.Year);
                char upper = char.ToUpper(str1[0]);
                str1 = string.Concat(upper.ToString(), str1.Substring(1));
                this._xFolderPath = string.Format("C:\\ALTOMOVUP\\{0}", str1);
                this.di = new DirectoryInfo(this._xFolderPath);

                if (!Directory.Exists("C:\\ALTOMOVUP" + "\\ArchivosAF"))
                {
                    // SE CREAN SUB CARPETAS PARA IDENTIFICAR LOS ARCHIVOS X6
                    this.di = Directory.CreateDirectory("C:\\ALTOMOVUP" + "\\ArchivosAF");
                    this.ltb_log.Items.Add(string.Format("{0} - Carpeta creada: {1}", DateTime.Now, "C:\\ALTOMOVUP" + "\\ArchivosAF"));
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }

                if (!Directory.Exists("C:\\ALTOMOVUP" + "\\ArchivosAFEnviado"))
                {
                    this.di = Directory.CreateDirectory("C:\\ALTOMOVUP" + "\\ArchivosAFEnviado");
                    this.ltb_log.Items.Add(string.Format("{0} - Carpeta creada: {1}", DateTime.Now, "C:\\ALTOMOVUP" + "\\ArchivosAFEnviado"));
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }
            }




            private void GetFiles()
            {
                string _fName = "";
                try
                {
                    DataTable dt = new DataTable();

                    SqlCommand sqlCommand = new SqlCommand("SP_Archivos214_20_AltoMovup", this._xConnString)
                    {
                        CommandType = CommandType.StoredProcedure
                    };

                    this._xConnString.Open();

                    SqlDataAdapter da = new SqlDataAdapter(sqlCommand);
                    da.Fill(dt);

                    this.ltb_log.Items.Add(string.Format("{0} - Transacciones Cargadas Correctamente..", DateTime.Now));
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        _fName = dt.Rows[i]["ARCHIVO"].ToString() + ".txt";
                        //Dado el caso, verifico que exista el archivo..
                        if (!File.Exists(Constantes.SourceDir + _fName))
                        {
                            File.Copy(Path.Combine(Constantes.SourceDir, _fName), Path.Combine("C:\\ALTOMOVUP" + "\\ArchivosAF", _fName), true);
                            this.ltb_log.Items.Add(string.Format("{0} - Archivo Procesado: {1}.", DateTime.Now, _fName));
                            this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                        }
                    }
                    this._xConnString.Close();
                }
                catch
                {
                    this._xConnString.Close();
                    //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                    ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                    this.ltb_log.Items.Add(string.Format("{0} - Error al cargar los Archivos.", DateTime.Now));
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }
            }
            private void UploadFiles()
            {

                string _fName = "";
                string str = "";
                string strFullNme = "";
                Session session = new Session();
                try
                {
                    // Setup session options
                    SessionOptions sessionOptions = new SessionOptions
                    {
                        Protocol = Protocol.Ftp,
                        FtpSecure= FtpSecure.Explicit,
                        FtpMode = FtpMode.Passive,
                        PortNumber = Constantes.PuertoSFTP,
                        HostName = Constantes.ServidorSFTP, //this.txtHostName.Text,
                        UserName = Constantes.UsuarioSFTP, //this.txtUserName.Text,
                        Password = Constantes.PassWordSFTP, //this.txtPwd.Text,
                        GiveUpSecurityAndAcceptAnySshHostKey = false,
                        GiveUpSecurityAndAcceptAnyTlsHostCertificate = true
                    };

                    session.Open(sessionOptions);
                    FileInfo[] files = (new DirectoryInfo(@"C:\ALTOMOVUP\\ArchivosAF")).GetFiles();
                    for (int i = 0; i < (int)files.Length; i++)
                    {
                        FileInfo fileInfo = files[i];
                        str = string.Concat(@"C:\ALTOMOVUP\ArchivosAFEnviado\", fileInfo.Name);
                        if ((fileInfo.Name.Contains(".txt") ? true : fileInfo.Name.Contains(".txt")))
                        {
                            String RemotePath = Constantes.RutaUploadSFTP; //this.txtRutaSubidaSFTP.Text; 
                            String filenameToUpload = fileInfo.Name;
                            String localPathtoUpload = @"C:\ALTOMOVUP\ArchivosAF\"; //this.txtRutaSubidaLocal.Text;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;
                            if (session.Opened)
                            {
                                //Get Ftp File
                                TransferOptions transferOptions = new TransferOptions();
                                transferOptions.TransferMode = TransferMode.Binary; //The Transfer Mode - 
                                                                                    //<em style="font-size: 9pt;">Automatic, Binary, or Ascii  
                                transferOptions.FilePermissions = null; //Permissions applied to remote files; 
                                                                        //null for default permissions.  Can set user, 
                                                                        //Group, or other Read/Write/Execute permissions. 
                                transferOptions.PreserveTimestamp = false; //Set last write time of 
                                                                           //destination file to that of source file - basically change the timestamp 
                                                                           //to match destination and source files.   
                                transferOptions.ResumeSupport.State = TransferResumeSupportState.Off;

                                TransferOperationResult transferResult;
                                //the parameter list is: local Path, Remote Path, Delete source file?, transfer Options  
                                transferResult = session.PutFiles(localPathtoUpload + filenameToUpload, RemotePath, false, transferOptions);
                                //Throw on any error 
                                transferResult.Check();
                                this.ltb_log.Items.Add(string.Format("{0} - Archivo enviado: {1}", DateTime.Now, fileInfo.Name));
                                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                            }
                            else
                            {
                                MessageBox.Show("Conexion error....", "FTP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                        ActualizaTransaccionEnviado(fileInfo.Name.Substring(7, 11));

                        if (!File.Exists(str))
                        {
                            File.Move(fileInfo.FullName, str);
                        }
                        else if (File.Exists(str))
                        {
                            File.Delete(fileInfo.FullName);
                        }
                    }

                    session.Close();
                    this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                    this.ltb_log.Items.Add("");
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }
                catch (Exception)
                {
                    //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                    ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                    if (!File.Exists(str))
                    {
                        File.Move(strFullNme, str);
                    }
                    else if (File.Exists(str))
                    {
                        File.Delete(strFullNme);
                    }
                    session.Close();
                    this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                    this.ltb_log.Items.Add("");
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                    //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            private void ActualizaTransaccionEnviado(string _xReferenceNumber)
            {
                try
                {
                    SqlCommand sqlCommand = new SqlCommand("SP_Actualiza_Transacciones_214_20", this._xConnString)
                    {
                        CommandType = CommandType.StoredProcedure
                    };

                    sqlCommand.Parameters.AddWithValue("@REFERENCENUMBER", _xReferenceNumber);
                    this._xConnString.Open();
                    sqlCommand.ExecuteNonQuery();
                    this._xConnString.Close();
                }
                catch
                {
                    this.ltb_log.Items.Add(string.Format("{0} - Problemas al cerrar el viaje.", DateTime.Now));
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }
            }
        #endregion

        #region Metodos_publicos
            public static Coordenates Convert2DegressMinutesSeconds(double position, bool isLong)
            {
                Coordenates ret = new Coordenates();

                // Negative: North
                // Positive: South
                // if isLong = true
                // Negative: East
                // Positive: West
                double absValue = Math.Abs(Math.Round(position * 1000000));
                int sign = Math.Sign(position);

                ret.Degress = (int)Math.Floor(absValue / 1000000);
                ret.Minutes = (int)Math.Floor(((absValue / 1000000) - Math.Floor(absValue / 1000000)) * 60);
                ret.Seconds = (Decimal)Math.Floor(((((absValue / 1000000) - Math.Floor(absValue / 1000000)) * 60) - Math.Floor(((absValue / 1000000) - Math.Floor(absValue / 1000000)) * 60)) * 100000) * 60 / 100000;

                if (isLong)
                    if (sign > 0)
                        ret.Geo = "W";
                    else
                        ret.Geo = "E";
                else
                    if (sign > 0)
                    ret.Geo = "N";
                else
                    ret.Geo = "S";

                return ret;
            }
            public static Coordenates Convert2DegressMinutesSeconds(double position)
            {
                // By default, it is a Latitude (North / South)
                return Convert2DegressMinutesSeconds(position, false);
            }
        #endregion

        #region Estructura
            public struct Coordenates
            {
                public int Degress;
                public int Minutes;
                public decimal Seconds;
                public String Geo; // N,S,E,W
            }
        #endregion
    }
}