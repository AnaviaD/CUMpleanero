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
    using System.Xml.Serialization;
    using System.Xml;
using FTRDHLFR;
#endregion

namespace CUMpleanero
{
    public partial class CUMpleanero : Form
    {
        #region Variables&Constantes
        private SqlConnection _xConnString = new SqlConnection();
        private SqlConnection _xConnString2 = new SqlConnection();
        private List<ArchivosDescarga> xTracksIDS = new List<ArchivosDescarga>();

        //private WebClient _xWebClient;
        //private ArchivosDescarga _ArchivosDescarga;
        //private string _xFolderPath;
        //private string _xNombreXML;
        //private string _xHTMLPass;
        private int _xManual;
        private int _xError;
        private string ErrorMessageTxt = "";
        #endregion

        #region Constructor
        public CUMpleanero()
        {
            //server = DESKTOP - FP59UDN\\SQLEXPRESS; database = x_FTR; Trusted_Connection = true
            //string _xServerName, string _xNombreBD, string _xTrustedConnection
            InitializeComponent();
            this.CargarConexion(Constantes._xServidorBD, Constantes._xNombreBD, Constantes._xUsuarioBD, Constantes._xPassWordBD);
            //this.CargarConexion2(Constantes._xServidorBD2, Constantes._xNombreBD2, Constantes._xUsuarioBD2, Constantes._xPassWordBD2);
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
                //this.ltb_log.Items.Add(string.Format("{0} - Error al actualizar la información, Intentelo Nuevamente...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                this.Cursor = Cursors.Default;
            }
        }
        #endregion

        //return Regex.Replace(strIn, @"[^\w\.@-]", " ", RegexOptions.None, TimeSpan.FromSeconds(1.5));
        static string CleanInput(string strIn)
        {
            // Replace invalid characters with empty strings.
            try
            {
                return Regex.Replace(strIn, "[^a-zA-Z0-9_.]+", " ", RegexOptions.Compiled, TimeSpan.FromSeconds(1.5));
            }

            catch (RegexMatchTimeoutException)
            {
                return String.Empty;
            }
        }

        #region Metodos Privados
        private void SFTP_Load(object sender, EventArgs e)
        {
            this.IniciarTimer();
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
                    //this.Proceso();
                    this._xManual = 0;
                    this.tmr_manual.Stop();
                    Thread.Sleep(400);
                    this.btn_actuaizar.Enabled = true;
                }
            }
            catch
            {
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                this.btn_actuaizar.Enabled = true;
            }
        }
        public void CargarConexion(string _xServerName, string _xNombreBD, string _xUser, string _xPassword)
        {
            //server = DESKTOP - FP59UDN\\SQLEXPRESS; database = x_FTR; Trusted_Connection = true
            try
            {
                this._xConnString.ConnectionString = string.Format("data source = {0}; initial catalog = {1}; User Id={2}; Password = {3};", new object[] { _xServerName, _xNombreBD, _xUser, _xPassword });
            }
            catch (Exception)
            {
                throw;
            }
        }
        private void CargarConexion2(string _xServerName, string _xNombreBD, string _xUser, string _xPassword)
        {
            this._xConnString2.ConnectionString = string.Format("data source = {0}; initial catalog = {1}; User Id={2}; Password = {3};", new object[] { _xServerName, _xNombreBD, _xUser, _xPassword });
        }
        private void IncreaseProgressBarPosiciones(object sender, EventArgs e)
        {
            if (this.pgb_Progreso.Value == this.pgb_Progreso.Maximum)
            {
                //this.Proceso();
                this.VerificaAlToque();
                this.pgb_Progreso.Value = 0;
                Thread.Sleep(200);
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
            //Tiempo de espera para refrescar el programa
            this.tmr_Actualiza.Interval = 300;
            this.tmr_Actualiza.Tick += new EventHandler(this.IncreaseProgressBarPosiciones);
        }
        private void Proceso()
        {
            try
            {
                System.Windows.Forms.Cursor cursor = this.Cursor;
                this.Cursor = Cursors.WaitCursor;
                this.ltb_log.Items.Add(string.Format("{0} - Sesión iniciada correctamente..", DateTime.Now));
                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                this._xError = 0;
                //this.DirectorioDescarga();
                this.pgb_Progreso.Value = 0;
                //this.GetFiles();
                this.pgb_Progreso.Value = 0;
                this.Cursor = cursor;
            }
            catch (Exception)
            {
                throw;
            }
        }
        #endregion



        public void VerificaAlToque()
        {


            try
            {
                //StartCheck
                HoyCumplesAnos hca = new HoyCumplesAnos();
                hca.StartCheck();
                this.ltb_log.Items.Add(string.Format("{0} - {1}", DateTime.Now, hca.ConsoladeSalida));
                this.ltb_log.Items.Add("");
            }
            catch (Exception)
            {

                throw;
            }




            //try
            //{
            //    // Setup session options

            //    bool allDirExists = System.IO.Directory.Exists(@"\\10.1.1.30\e$\Attachments");

            //    if (allDirExists)
            //    {

            //        var amazonDirectories = @"\\10.1.1.30\e$\Attachments\sftr-amazon";
            //        //var testDirectories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";


            //        bool amazonDirexists = System.IO.Directory.Exists(amazonDirectories);
            //        //bool testDirexists = System.IO.Directory.Exists(testDirectories);

            //        //if (testDirexists)
            //        //{
            //        //    FileInfo[] testFiles = (new DirectoryInfo(testDirectories)).GetFiles();
            //        //    if (testFiles.Length > 0)
            //        //        GetXml_Files();
            //        //}

            //        if (amazonDirexists)
            //        {
            //            FileInfo[] amazonFiles = (new DirectoryInfo(amazonDirectories)).GetFiles();
            //            if (amazonFiles.Length > 0)
            //                Console.WriteLine("Hey hola");
            //            this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
            //            this.ltb_log.Items.Add("");
            //        }

            //    }


            //    this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
            //    this.ltb_log.Items.Add("");
            //    //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            //}
            //catch (Exception ex)
            //{
            //    string erro = ex.ToString();
            //    //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
            //    //ActualizaTransaccionEnviado(_fName.Substring(7, 11));

            //    this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Leer Carpetas...", erro));
            //    this.ltb_log.Items.Add("");
            //    //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            //    //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}

        }
    }

}