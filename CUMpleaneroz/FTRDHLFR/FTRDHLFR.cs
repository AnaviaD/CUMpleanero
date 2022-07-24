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
#endregion

namespace FTRDHLFR
{
    public partial class FTRDHLFR : Form
    {
        #region Variables&Constantes
        private SqlConnection _xConnString = new SqlConnection();
        private SqlConnection _xConnString2 = new SqlConnection();
        private DirectoryInfo di;
        private List<string> xRenglon = new List<string>();
        private List<string> BannedExcel = new List<string>();
        private List<ArchivosDescarga> xTracksIDS = new List<ArchivosDescarga>();

        //private WebClient _xWebClient;
        //private ArchivosDescarga _ArchivosDescarga;
        private string _xFolderPath;
        //private string _xNombreXML;
        //private string _xHTMLUser;
        //private string _xHTMLPass;
        private int _xManual;
        private int _xError;
        private string ErrorMessageTxt = "";
        #endregion

        #region Constructor
        public FTRDHLFR()
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
            if (this.pgb_Progreso.Value == this.pgb_Progreso.Maximum)
            {
                //this.Proceso();
                this.VerificaAlToque();
                this.VerificaPendientes();
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


        /// <summary>
        /// Elimina los Materiales previos de un cliente 
        /// </summary>
        /// <param name="referenciaDelServicio">Identificador</param>
        /// <param name="rFCdelRemitente">RFC</param>
        private void EliminaMaterPrevios(string referenciaDelServicio, string rFCdelRemitente)
        {
            SqlCommand sqlCommand = new SqlCommand("USP_DelCFDCartaPorteMatProveRecolectas", this._xConnString)
            {
                CommandType = CommandType.StoredProcedure
            };

            sqlCommand.Parameters.AddWithValue("@codigo", referenciaDelServicio.Trim());
            sqlCommand.Parameters.AddWithValue("@RFC", rFCdelRemitente.Trim());

            this._xConnString.Open();
            sqlCommand.ExecuteNonQuery();
            this._xConnString.Close();
        }

        private void EliminaMaterPreviosStellantis(List<EntityStellantis> StellantisField)
        {
            foreach (var r in StellantisField)
            { 
                SqlCommand sqlCommand = new SqlCommand("USP_DelCFDCartaPorteMatProveRecolectasStellantis", this._xConnString)
                {
                    CommandType = CommandType.StoredProcedure
                };

                sqlCommand.Parameters.AddWithValue("@UniqueRecordIdentifier", r.UniqueRecordIdentifier.Trim());
                sqlCommand.Parameters.AddWithValue("@MerchandiseOwnerTaxID", r.MerchandiseOwnerTaxID.Trim());

                this._xConnString.Open();
                sqlCommand.ExecuteNonQuery();
                this._xConnString.Close();
            
            }
        }


        private List<EntityStellantis> VerificaMercanciasStellantis(List<EntityStellantis> StellantisList)
        {
            List<EntityStellantis> StellantCopy = new List<EntityStellantis>();
            int i = 0;
            foreach (var r in StellantisList)
            {
                try
                {
                    DataTable dt = new DataTable();
                    //string sentencia = string.Format("SELECT TOP 1 * FROM TB_ASIGNACION_LOADS WHERE Ruta in('{0}') AND DATEDIFF(DAY,Fecha,GETDATE()) <= 0 ORDER BY Fecha ", Ruta);
                    string sentencia = string.Format("SELECT* FROM x_FTR.dbo.TB_MERCANCIAS_STELLANTIS WHERE UniqueRecordIdentifier = '{0}'", r.UniqueRecordIdentifier);
                    //SELECT* FROM x_FTR.dbo.TB_MERCANCIAS_STELLANTIS WHERE UniqueRecordIdentifier = '2022XFTR000001278428'
                    this._xConnString.Open();

                    SqlDataAdapter dataAda = new SqlDataAdapter(sentencia, this._xConnString);
                    dataAda.Fill(dt);
                    this._xConnString.Close();

                    if (dt != null)
                    {
                        if (dt.Rows.Count > 0)
                        {
                            //StellantCopy.RemoveAt(i);
                        }else { StellantCopy.Add(r); }
                    }
                }
                catch
                {
                    this.ltb_log.Items.Add(string.Format("{0} - Problemas al obtener catalogos los archivos.", DateTime.Now));
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }
                i++;
            }
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", DateTime.Now));
            return StellantCopy;
        }

        private List<EntintDHLFields> VerificaMercanciasClientes(List<EntintDHLFields> StellantisList)
        {
            List<EntintDHLFields> ClientCopy = new List<EntintDHLFields>();
            foreach (var r in StellantisList)
            {
                try
                {
                    DataTable dt = new DataTable();
                    string sentencia = string.Format("SELECT TOP (1) * FROM x_FTR.dbo.TB_MERCANCIAS_CLIENTE WHERE ReferenciaDelServicio = '{0}'", r.ReferenciaDelServicio);
                    this._xConnString.Open();

                    SqlDataAdapter dataAda = new SqlDataAdapter(sentencia, this._xConnString);
                    dataAda.Fill(dt);
                    this._xConnString.Close();

                    if (dt != null)
                    {
                        if (dt.Rows.Count > 0)
                        {
                            //ClientCopy.RemoveAt(i);
                        }
                        else { ClientCopy.Add(r); }
                    }
                }
                catch
                {
                    this.ltb_log.Items.Add(string.Format("{0} - Problemas al obtener catalogos los archivos.", DateTime.Now));
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }
            }
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", DateTime.Now));
            return ClientCopy;
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


        private void InsertaMercanciaClientesList(List<EntintDHLFields> fieldItem, string directories, string fileName)
        {
            Regex regexObj = new Regex(@"[^\d]");
            Regex regexObjdec = new Regex(@"^[0-9.-]+$");
            string Municipiofmt = "000";
            string CPfmt = "00000";

            foreach (var r in fieldItem)
            {
                try
                {
                    
                    r.CP = String.IsNullOrEmpty(r.CP) ? "00000" : r.CP;
                    r.CP2 = String.IsNullOrEmpty(r.CP2) ? "00000" : r.CP2;
                    
                    if (String.IsNullOrEmpty(r.Municipio)) { } else { r.Municipio = Convert.ToInt32(regexObj.Replace(r.Municipio.Replace(" ", ""), "")).ToString(Municipiofmt); }
                    if (String.IsNullOrEmpty(r.Municipio2)) { } else { r.Municipio2 = Convert.ToInt32(regexObj.Replace(r.Municipio2.Replace(" ", ""), "")).ToString(Municipiofmt); }

                    r.RFCdelRemitente = String.IsNullOrEmpty(r.RFCdelRemitente) ? "XAXX010101000" : r.RFCdelRemitente;
                    r.RFCDestinatario = String.IsNullOrEmpty(r.RFCDestinatario) ? "XAXX010101000" : r.RFCDestinatario;

                    r.PesoNeto = String.IsNullOrEmpty(r.PesoNeto) ? "1" : r.PesoNeto;
                    r.PesoNeto = String.Equals(r.PesoNeto, "0") ? "1" : r.PesoNeto;
                    r.PesoNeto = regexObjdec.Match(r.PesoNeto.Replace(" ", "")).ToString();
                    r.PesoNeto = String.Format("{0:0.##}", r.PesoNeto);

                    r.RFCdelRemitente = r.RFCdelRemitente.Replace(" ", "").ToUpper();
                    r.RFCDestinatario = r.RFCDestinatario.Replace(" ", "").ToUpper();
                    r.CP = Convert.ToInt32(regexObj.Replace(r.CP.Replace(" ", ""), "")).ToString(CPfmt);
                    r.CP2 = Convert.ToInt32(regexObj.Replace(r.CP2.Replace(" ", ""), "")).ToString(CPfmt);
                    r.ValorDeLaMercancia = regexObj.Replace(r.ValorDeLaMercancia.Replace(" ", ""), "");
                    r.ValorDeLaMercancia = String.IsNullOrEmpty(r.ValorDeLaMercancia) ? "1" : r.ValorDeLaMercancia;
                    r.ValorDeLaMercancia = String.Equals(r.ValorDeLaMercancia, "0") ? "1" : r.ValorDeLaMercancia;
                    r.NumeroTotalMercancias = String.IsNullOrEmpty(r.NumeroTotalMercancias) ? "1" : r.NumeroTotalMercancias;
                    r.NumeroTotalMercancias = regexObj.Replace(r.NumeroTotalMercancias.Replace(" ", ""), "");
                    r.NumeroTotalMercancias = String.Equals(r.NumeroTotalMercancias, "0") ? "1" : r.NumeroTotalMercancias;
                    r.MaterialPeligroso = String.IsNullOrEmpty(r.MaterialPeligroso) ? "No" : r.MaterialPeligroso;
                    r.MaterialPeligroso = String.Equals(r.MaterialPeligroso, "NO") ? "No" : r.MaterialPeligroso;
                    r.TipoDeMoneda = String.IsNullOrEmpty(r.TipoDeMoneda) ? "MXN" : r.TipoDeMoneda;
                }
                catch (Exception)
                {
                    ErrorMessageTxt += String.Format("Puede que Los campos CP, Peso Total, Valor Total, Numero de Mercancías, contengan caracteres{0}", "<br />");
                    EnviarEmail(directories, fileName, ErrorMessageTxt);
                    throw;
                }


                try
                {

                    SqlCommand sqlCommand = new SqlCommand("SP_InsertaMercanciaClientes", this._xConnString)
                    {
                        CommandType = CommandType.StoredProcedure
                    };

                    sqlCommand.Parameters.AddWithValue("@ReferenciaDelServicio", r.ReferenciaDelServicio.Trim());
                    sqlCommand.Parameters.AddWithValue("@RSdelRemitente", r.RSdelRemitente.Trim());
                    sqlCommand.Parameters.AddWithValue("@RFCdelRemitente", r.RFCdelRemitente.Trim());
                    sqlCommand.Parameters.AddWithValue("@Supplier", r.Supplier.Trim());
                    sqlCommand.Parameters.AddWithValue("@Calle", r.Calle.Trim());
                    sqlCommand.Parameters.AddWithValue("@Municipio", r.Municipio.Trim());
                    sqlCommand.Parameters.AddWithValue("@Estado", r.Estado.Trim());
                    sqlCommand.Parameters.AddWithValue("@Pais", r.Pais.Trim());
                    sqlCommand.Parameters.AddWithValue("@CP", r.CP.Trim());
                    sqlCommand.Parameters.AddWithValue("@RSdelDestinatario", r.RSdelDestinatario.Trim());
                    sqlCommand.Parameters.AddWithValue("@RFCDestinatario", r.RFCDestinatario.Trim());
                    sqlCommand.Parameters.AddWithValue("@Calle2", r.Calle2.Trim());
                    sqlCommand.Parameters.AddWithValue("@Municipio2", r.Municipio2.Trim());
                    sqlCommand.Parameters.AddWithValue("@Estado2", r.Estado2.Trim());
                    sqlCommand.Parameters.AddWithValue("@Pais2", r.Pais2.Trim());
                    sqlCommand.Parameters.AddWithValue("@CP2", r.CP2.Trim());
                    sqlCommand.Parameters.AddWithValue("@PesoNeto", r.PesoNeto.Trim());
                    sqlCommand.Parameters.AddWithValue("@NumeroTotalMercancias", r.NumeroTotalMercancias.Trim());
                    sqlCommand.Parameters.AddWithValue("@ClaveDelBienTransportado", r.ClaveDelBienTransportado.Trim());
                    sqlCommand.Parameters.AddWithValue("@DescripcionDelBienTransportado", r.DescripcionDelBienTransportado.Trim());
                    sqlCommand.Parameters.AddWithValue("@ClaveUnidadDeMedida", r.ClaveUnidadDeMedida.Trim());
                    sqlCommand.Parameters.AddWithValue("@MaterialPeligroso", r.MaterialPeligroso.Trim());
                    sqlCommand.Parameters.AddWithValue("@ValorDeLaMercancia", r.ValorDeLaMercancia.Trim());
                    sqlCommand.Parameters.AddWithValue("@TipoDeMoneda", r.TipoDeMoneda.Trim());

                    this._xConnString.Open();
                    sqlCommand.ExecuteNonQuery();
                    this._xConnString.Close();

                }
                catch(Exception ex)
                {
                    this.ltb_log.Items.Add(string.Format("{0} {1}- Problemas al insertar los archivos.", DateTime.Now, ex.Message));
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }
            }
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", DateTime.Now));
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", fileName));
        }


        private void InsertaMercanciaClientesStellantisList(List<EntityStellantis> fieldItem, string directories, string fileName)
        {
            
            foreach (var r in fieldItem)
            {

                try
                {

                    SqlCommand sqlCommand = new SqlCommand("SP_InsertaMercanciaStellantis", this._xConnString)
                    {
                        CommandType = CommandType.StoredProcedure
                    };

                    sqlCommand.Parameters.AddWithValue("@UniqueRecordIdentifier", r.UniqueRecordIdentifier.Trim());
                    sqlCommand.Parameters.AddWithValue("@SenderID", r.SenderID.Trim());
                    sqlCommand.Parameters.AddWithValue("@ReciverID", r.ReciverID.Trim());
                    sqlCommand.Parameters.AddWithValue("@MerchandiseOwnerTaxID", r.MerchandiseOwnerTaxID.Trim());
                    sqlCommand.Parameters.AddWithValue("@MerchandiseOwner", r.MerchandiseOwner.Trim());
                    sqlCommand.Parameters.AddWithValue("@MerchandiseOwnerAddressLine1", r.MerchandiseOwnerAddressLine1.Trim());
                    sqlCommand.Parameters.AddWithValue("@MerchandiseOwnerAddressLine2", r.MerchandiseOwnerAddressLine2.Trim());
                    sqlCommand.Parameters.AddWithValue("@MerchandiseOwnerCity", r.MerchandiseOwnerCity.Trim());
                    sqlCommand.Parameters.AddWithValue("@MerchandiseOwnerState", r.MerchandiseOwnerState.Trim());
                    sqlCommand.Parameters.AddWithValue("@MerchandiseOwnerZip", r.MerchandiseOwnerZip.Trim());
                    sqlCommand.Parameters.AddWithValue("@MerchandiseMunicipalySATreference", r.MerchandiseMunicipalySATreference.Trim());
                    sqlCommand.Parameters.AddWithValue("@MerchandiseOwnerCntry", r.MerchandiseOwnerCntry.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipTo", r.ShipTo.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipToName", r.ShipToName.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipToAddressLine", r.ShipToAddressLine.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipToCity", r.ShipToCity.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipToState", r.ShipToState.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipToZip", r.ShipToZip.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipToMunicipalySATreference", r.ShipToMunicipalySATreference.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipToCntry", r.ShipToCntry.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipFrom", r.ShipFrom.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierTaxIdentifier", r.SupplierTaxIdentifier.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierName", r.SupplierName.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierAddressLine", r.SupplierAddressLine.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierCity", r.SupplierCity.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierState", r.SupplierState.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierZip", r.SupplierZip.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierMuniciplalitySATreference", r.SupplierMuniciplalitySATreference.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierCntry", r.SupplierCntry.Trim());
                    sqlCommand.Parameters.AddWithValue("@PartOrContainerID", r.PartOrContainerID.Trim());
                    sqlCommand.Parameters.AddWithValue("@PartDescription", r.PartDescription.Trim());
                    sqlCommand.Parameters.AddWithValue("@PartSATCode", r.PartSATCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@PartSATDescription", r.PartSATDescription.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShippedQuantity", r.ShippedQuantity.Trim());
                    sqlCommand.Parameters.AddWithValue("@UnitOfMeasureShipped", r.UnitOfMeasureShipped.Trim());
                    sqlCommand.Parameters.AddWithValue("@UnitOfMeasureSATCode", r.UnitOfMeasureSATCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@UnitOfMeasureSATdescription", r.UnitOfMeasureSATdescription.Trim());
                    sqlCommand.Parameters.AddWithValue("@HazmatFlag", r.HazmatFlag.Trim());
                    sqlCommand.Parameters.AddWithValue("@HazmatSATCode", r.HazmatSATCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@HazmatSATDescription", r.HazmatSATDescription.Trim());
                    sqlCommand.Parameters.AddWithValue("@ContainerIdentifier", r.ContainerIdentifier.Trim());
                    sqlCommand.Parameters.AddWithValue("@I_SAT_CNTNR", r.I_SAT_CNTNR.Trim());
                    sqlCommand.Parameters.AddWithValue("@ContainerSATDescription", r.ContainerSATDescription.Trim());
                    sqlCommand.Parameters.AddWithValue("@ContainerQty", r.ContainerQty.Trim());
                    sqlCommand.Parameters.AddWithValue("@ContainerTareWeight", r.ContainerTareWeight.Trim());
                    sqlCommand.Parameters.AddWithValue("@NetShipmentWeight", r.NetShipmentWeight.Trim());
                    sqlCommand.Parameters.AddWithValue("@GrossShipmentWeight", r.GrossShipmentWeight.Trim());
                    sqlCommand.Parameters.AddWithValue("@UnitOfMeasureWeight", r.UnitOfMeasureWeight.Trim());
                    sqlCommand.Parameters.AddWithValue("@HTSCode", r.HTSCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@HTSCountryCode", r.HTSCountryCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@CurrencyCode", r.CurrencyCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierCode", r.SupplierCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@FinalDestinationCode", r.FinalDestinationCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipmentIdentifier", r.ShipmentIdentifier.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierPackingSlip", r.SupplierPackingSlip.Trim());
                    sqlCommand.Parameters.AddWithValue("@SupplierBillOfLoading", r.SupplierBillOfLoading.Trim());
                    sqlCommand.Parameters.AddWithValue("@FreightConsolidationBillOfLandingNumber", r.FreightConsolidationBillOfLandingNumber.Trim());
                    sqlCommand.Parameters.AddWithValue("@ConsolidationShipmentIdentifier", r.ConsolidationShipmentIdentifier.Trim());
                    sqlCommand.Parameters.AddWithValue("@PoolpointShipfrom", r.PoolpointShipfrom.Trim());
                    sqlCommand.Parameters.AddWithValue("@PoolPointShipto", r.PoolPointShipto.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipmentDate", r.ShipmentDate.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipmentTime", r.ShipmentTime.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipmentTimestamp", r.ShipmentTimestamp.Trim());
                    sqlCommand.Parameters.AddWithValue("@CarrierSCAC", r.CarrierSCAC.Trim());
                    sqlCommand.Parameters.AddWithValue("@ConveyanceIdentifier", r.ConveyanceIdentifier.Trim());
                    sqlCommand.Parameters.AddWithValue("@OwnerSCAC", r.OwnerSCAC.Trim());
                    sqlCommand.Parameters.AddWithValue("@TransportationMode", r.TransportationMode.Trim());
                    sqlCommand.Parameters.AddWithValue("@PartCountInContainer", r.PartCountInContainer.Trim());
                    sqlCommand.Parameters.AddWithValue("@AETCNumer", r.AETCNumer.Trim());
                    sqlCommand.Parameters.AddWithValue("@LotNumber", r.LotNumber.Trim());
                    sqlCommand.Parameters.AddWithValue("@ChampsTransactionCode", r.ChampsTransactionCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@ChampsPurposeCode", r.ChampsPurposeCode.Trim());
                    sqlCommand.Parameters.AddWithValue("@ASNStatus", r.ASNStatus.Trim());
                    sqlCommand.Parameters.AddWithValue("@ShipmentIdentifierCount", r.ShipmentIdentifierCount.Trim());
                    sqlCommand.Parameters.AddWithValue("@ASNCount", r.ASNCount.Trim());
                    sqlCommand.Parameters.AddWithValue("@MasterBilOfLadinng", r.MasterBilOfLadinng.Trim());
                    sqlCommand.Parameters.AddWithValue("@UnitEstimatedCost", r.UnitEstimatedCost.Trim());
                    sqlCommand.Parameters.AddWithValue("@FillerForFutureUser", r.FillerForFutureUser.Trim());
                    sqlCommand.Parameters.AddWithValue("@TotalWeightContainer", r.TotalWeightContainer.Trim());
                    sqlCommand.Parameters.AddWithValue("@TotalWeightContainerUnit", r.TotalWeightContainerUnit.Trim());
                    sqlCommand.Parameters.AddWithValue("@GrossShipmentWeightUnit", r.GrossShipmentWeightUnit.Trim());

                    this._xConnString.Open();
                    sqlCommand.ExecuteNonQuery();
                    this._xConnString.Close();

                }
                catch (Exception ex)
                {
                    this.ltb_log.Items.Add(string.Format("{0} {1}- Problemas al insertar los archivos.", DateTime.Now, ex.Message));
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }
            }
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", DateTime.Now));
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", fileName));
        }

        //EXEC X_FTR.DBO.USP_GetOrigenCarga 28358-- Lunes
        //EXEC X_FTR.DBO.USP_GetCatalogo_Referencia_SAT_CP '98519'

        //EXEC X_FTR.DBO.USP_GetDestinoCarga 28358
        //EXEC X_FTR.DBO.USP_GetCatalogo_Referencia_SAT_CP 36118

        //EXEC X_FTR.DBO.USP_GetOrigenCarga 28224-- Miércoles
        //EXEC X_FTR.DBO.USP_GetCatalogo_Referencia_SAT_CP '52960'

        //EXEC X_FTR.DBO.USP_GetDestinoCarga 28224
        //EXEC X_FTR.DBO.USP_GetCatalogo_Referencia_SAT_CP 71720

        //-- sftr - learMexico   Directorio

        //-- sftr-learMexico Directorio

        //SELECT* FROM SAT_CP WHERE Estado='OAX' AND CP >='70000'

        //SELECT* FROM x_ftr.dbo.TB_CODIGO_CARGA_PENSKE WHERE Numero_Ruta ='AQ049' ORDER BY Fecha_Inicio_Viaje
        //SELECT* FROM x_ftr.dbo.TB_CODIGO_CARGA_PENSKE WHERE Numero_Ruta ='JG032' ORDER BY Fecha_Inicio_Viaje
        //SELECT* FROM TB_ASIGNACION_LOADS WHERE Ruta in('AQ049', 'JG032') ORDER BY Fecha
        //SELECT* FROM truenvio where envio_idno in ('28358','28224')

        //EXEC X_FTR.DBO.USP_GetCatalogo_Referencia_SAT_CP 36275;

        private DataTable ObtenerCargaPenske(string Ruta)
        {
            DataTable dt = new DataTable();
            try
            {
                string sentencia = string.Format("SELECT TOP 1 * FROM x_ftr.dbo.TB_ASIGNACION_LOADS WHERE Codigo_Carga in('{0}')", Ruta);
                this._xConnString.Open();

                SqlDataAdapter dataAda = new SqlDataAdapter(sentencia, this._xConnString);
                dataAda.Fill(dt);
                this._xConnString.Close();
                return dt;
            }
            catch
            {
                this.ltb_log.Items.Add(string.Format("{0} - Problemas al obtener catalogos los archivos.", DateTime.Now));
                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", DateTime.Now));
            return dt;
        }

        private DataTable ObtenerOrigenCarga(int codigoVehiculo)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlCommand sqlCommand = new SqlCommand("[USP_GetOrigenCarga]", this._xConnString)
                {
                    CommandType = CommandType.StoredProcedure
                };

                sqlCommand.Parameters.AddWithValue("@UE", codigoVehiculo);

                this._xConnString.Open();
                SqlDataReader dataTViajes = sqlCommand.ExecuteReader();
                dt.Load(dataTViajes);
                this._xConnString.Close();
                return dt;
            }
            catch
            {
                this.ltb_log.Items.Add(string.Format("{0} - Problemas al obtener catalogos los archivos.", DateTime.Now));
                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", DateTime.Now));
            return dt;
        }


        private DataTable ObtenerDestinoCarga(int codigoVehiculo)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlCommand sqlCommand = new SqlCommand("[USP_GetDestinoCarga]", this._xConnString)
                {
                    CommandType = CommandType.StoredProcedure
                };

                sqlCommand.Parameters.AddWithValue("@UE", codigoVehiculo);

                this._xConnString.Open();
                SqlDataReader dataTViajes = sqlCommand.ExecuteReader();
                dt.Load(dataTViajes);
                this._xConnString.Close();
                return dt;
            }
            catch
            {
                this.ltb_log.Items.Add(string.Format("{0} - Problemas al obtener catalogos los archivos.", DateTime.Now));
                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", DateTime.Now));
            return dt;
        }


        //SqlCommand sqlCommand = new SqlCommand("[SAT_CP]", this._xConnString) tabla, no sp
        //SqlCommand sqlCommand = new SqlCommand("[USP_GetCatalogo_Referencia_SAT_CP]", this._xConnString)
        //SqlCommand sqlCommand = new SqlCommand("[[USP_GetCatalogo_Colonia]]", this._xConnString)
        //SqlCommand sqlCommand = new SqlCommand("[[USP_GetCatalogo_Estados]]", this._xConnString)
        private DataTable ObtenerReferenciaSat(string CodP)
        {
            DataTable dt = new DataTable();
            try
            {
                SqlCommand sqlCommand = new SqlCommand("[USP_GetCatalogo_Referencia_SAT_CP]", this._xConnString)
                {
                    CommandType = CommandType.StoredProcedure
                };

                sqlCommand.Parameters.AddWithValue("@CodigoPostal", CodP);

                this._xConnString.Open();
                SqlDataReader dataTViajes = sqlCommand.ExecuteReader();
                dt.Load(dataTViajes);
                this._xConnString.Close();
                if (dt.Rows.Count == 0)
                {

                    DataRow dr;
                    dr = dt.NewRow();
                    dr[0] = "0";
                    dr[1] = "0";
                    dr[2] = "0";
                    dr[3] = "0";
                    dt.Rows.Add(dr);
                }
                return dt;
            }
            catch
            {
                this.ltb_log.Items.Add(string.Format("{0} - Problemas al obtener catalogos los archivos.", DateTime.Now));
                this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            this.ltb_log.Items.Add(string.Format("{0} - Archivo Completado...", DateTime.Now));
            return dt;
        }

        //Private Sub VerificaPeticiones(lista As List(Of String), ByVal directorio As String)
        //Dim di As New IO.DirectoryInfo(directorio)
        //Dim diar1 As IO.FileInfo() = di.GetFiles("*.lgs")
        //Dim dra As IO.FileInfo

        //For Each dra In diar1
        //    If (dra.Name.Contains(".lgs")) Then
        //        lista.Add(dra.FullName.ToString())
        //    End If
        //Next
        //End Sub

        private void VerificaPeticiones(List<string> lista, string directorio)
        {
            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(directorio);
            System.IO.FileInfo[] diar1 = di.GetFiles("*.xlm");
            //System.IO.FileInfo dirFileInf;

            foreach (var dirFileInf in diar1)
            {
                if ((dirFileInf.Name.Contains(".xlm")))
                    lista.Add(dirFileInf.FullName.ToString());
            }
        }

        private void VerificaPendientes()
        {
            for (int i = 0; i < BannedExcel.Count; i++)
            {
                FileStream fileStream = null;

                // limpieza ante todo

                FileInfo datos = new FileInfo(BannedExcel[i]);
                ExcelPackage LecExcel = new ExcelPackage(datos);
                try
                {
                    fileStream = datos.Open(FileMode.Open, FileAccess.ReadWrite);
                    BannedExcel.Remove(BannedExcel[i]);
                }
                catch (Exception)
                {
                }

            }
            
        }

        string EstaOnoBaneado(string filDir)
        {
            for (int i = 0; i < BannedExcel.Count; i++)
            {
                if (BannedExcel[i] == filDir)
                {
                    return "SI";
                }
            }
            return "NO";
        }

        private void VerificaAlToque()
        {

            try
            {
                // Setup session options

                bool allDirExists = System.IO.Directory.Exists(@"\\10.1.1.30\e$\Attachments");

                if (allDirExists)
                {

                    var amazonDirectories = @"\\10.1.1.30\e$\Attachments\sftr-amazon";
                    var androidDirectories = @"\\10.1.1.30\e$\Attachments\sftr-android";
                    var aplDirectories = @"\\10.1.1.30\e$\Attachments\sftr-apl";
                    var celticDirectories = @"\\10.1.1.30\e$\Attachments\sftr-celtic";
                    var dhlDirectories = @"\\10.1.1.30\e$\Attachments\sftr-dhlsupply";
                    var genericDirectories = @"\\10.1.1.30\e$\Attachments\sftr-generico";
                    var jbhuntDirectories = @"\\10.1.1.30\e$\Attachments\sftr-jbhunt";
                    var learDirectories = @"\\10.1.1.30\e$\Attachments\sftr-learmexico";
                    var trupperDirectories = @"\\10.1.1.30\e$\Attachments\sftr-truper";
                    var tremecDirectories = @"\\10.1.1.30\e$\Attachments\sftr-tremec";
                    var transplaceDirectories = @"\\10.1.1.30\e$\Attachments\sftr-transplace";
                    var stellantisDirectories = @"\\10.1.1.30\e$\Attachments\sftr-stellantis";
                    var stellantisMTsDirectories = @"\\10.1.1.30\e$\Attachments\sftr-stellantis-mts";
                    var schneiderDirectories = @"\\10.1.1.30\e$\Attachments\sftr-schneider";
                    //var testDirectories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";


                    bool amazonDirexists = System.IO.Directory.Exists(amazonDirectories);
                    bool androidDirexists = System.IO.Directory.Exists(androidDirectories);
                    bool aplDirexists = System.IO.Directory.Exists(aplDirectories);
                    bool celticDirexists = System.IO.Directory.Exists(celticDirectories);
                    bool dhlDirexists = System.IO.Directory.Exists(dhlDirectories);
                    bool genericDirexists = System.IO.Directory.Exists(genericDirectories);
                    bool jbhuntDirexists = System.IO.Directory.Exists(jbhuntDirectories);
                    bool learDirexists = System.IO.Directory.Exists(learDirectories);
                    bool trupperDirexists = System.IO.Directory.Exists(trupperDirectories);
                    bool tremecDirexists = System.IO.Directory.Exists(tremecDirectories);
                    bool transplaceDirexists = System.IO.Directory.Exists(transplaceDirectories);
                    bool stellantisDirexists = System.IO.Directory.Exists(stellantisDirectories);
                    bool stellantisMTsDirexists = System.IO.Directory.Exists(stellantisMTsDirectories);
                    bool schneiderDirexists = System.IO.Directory.Exists(schneiderDirectories);
                    //bool testDirexists = System.IO.Directory.Exists(testDirectories);

                    //if (testDirexists)
                    //{
                    //    FileInfo[] testFiles = (new DirectoryInfo(testDirectories)).GetFiles();
                    //    if (testFiles.Length > 0)
                    //        GetXml_Files();
                    //}

                    if (amazonDirexists)
                    {
                        FileInfo[] amazonFiles = (new DirectoryInfo(amazonDirectories)).GetFiles();
                        if (amazonFiles.Length > 0)
                            GetAmazon_Files();
                    }

                    if (androidDirexists)
                    {
                        FileInfo[] androidFiles = (new DirectoryInfo(androidDirectories)).GetFiles();
                        if (androidFiles.Length > 0)
                            GetAndroid_Files();
                    }

                    if (aplDirexists)
                    {
                        FileInfo[] aplFiles = (new DirectoryInfo(aplDirectories)).GetFiles();
                        if (aplFiles.Length > 0)
                            GetAPL_Files();
                    }

                    if (celticDirexists)
                    {
                        FileInfo[] celticFiles = (new DirectoryInfo(celticDirectories)).GetFiles();
                        if (celticFiles.Length > 0)
                            GetCeltic_Files();
                    }

                    if (dhlDirexists)
                    {
                        FileInfo[] dhlFiles = (new DirectoryInfo(dhlDirectories)).GetFiles();
                        if (dhlFiles.Length > 0)
                            GetDHLAttachment_Files();
                    }

                    if (genericDirexists)
                    {
                        FileInfo[] genericFiles = (new DirectoryInfo(genericDirectories)).GetFiles();
                        if (genericFiles.Length > 0)
                            GetGeneric_Files();
                    }

                    if (jbhuntDirexists)
                    {
                        FileInfo[] jbhuntFiles = (new DirectoryInfo(jbhuntDirectories)).GetFiles();
                        if (jbhuntFiles.Length > 0)
                            GetJBHunt_Files();
                    }

                    if (learDirexists)
                    {
                        FileInfo[] learFiles = (new DirectoryInfo(learDirectories)).GetFiles();
                        if (learFiles.Length > 0)
                            GetLear_Files();
                    }

                    if (trupperDirexists)
                    {
                        FileInfo[] trupperFiles = (new DirectoryInfo(trupperDirectories)).GetFiles();
                        if (trupperFiles.Length > 0)
                            GetTruper_Files();
                    }

                    if (tremecDirexists)
                    {
                        FileInfo[] tremecFiles = (new DirectoryInfo(tremecDirectories)).GetFiles();
                        if (tremecFiles.Length > 0)
                            GetTremec_Files();
                    }

                    if (transplaceDirexists)
                    {
                        FileInfo[] transplFiles = (new DirectoryInfo(transplaceDirectories)).GetFiles();
                        if (transplFiles.Length > 0)
                            GetXml_Files();
                    }

                    if (stellantisDirexists)
                    {
                        FileInfo[] stellantislFiles = (new DirectoryInfo(stellantisDirectories)).GetFiles();
                        if (stellantislFiles.Length > 0)
                            GetStellantis_Files();
                    }

                    if (stellantisMTsDirexists)
                    {
                        FileInfo[] stellantisMTSFiles = (new DirectoryInfo(stellantisMTsDirectories)).GetFiles();
                        if (stellantisMTSFiles.Length > 0)
                            GetStellantisMTS_Files();
                    }

                    if (schneiderDirexists)
                    {
                        FileInfo[] schneidrMTSFiles = (new DirectoryInfo(schneiderDirectories)).GetFiles();
                        if (schneidrMTSFiles.Length > 0)
                            Getschneider_Files();
                    }

                }


                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));

                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Leer Carpetas...", erro));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void GetClientFiles_Message()
        {

            try
            {
                // Setup session options
                

                bool dhlexists = System.IO.Directory.Exists(@"\\10.1.1.30\e$\Attachments");

                if (dhlexists)
                {
                    var Attachdirectories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments");
                    for (int i = 1; i < (int)Attachdirectories.Length; i++)
                    {
                        if ((Attachdirectories[i].Contains("sftr-") ? true : Attachdirectories [i].Contains("sftr-")))
                        {
                            if ((Attachdirectories [i].Contains("android") ? true : Attachdirectories [i].Contains("android")))
                            {
                                //GetXml_Files();
                                GetAndroid_Files();
                            }
                            if ((Attachdirectories [i].Contains("celtic") ? true : Attachdirectories [i].Contains("celtic")))
                            {
                                GetCeltic_Files();
                            }
                            if ((Attachdirectories [i].Contains("amazon") ? true : Attachdirectories [i].Contains("amazon")))
                            {
                                GetAmazon_Files();
                            }
                            if ((Attachdirectories [i].Contains("generico") ? true : Attachdirectories [i].Contains("generico")))
                            {
                                GetGeneric_Files();
                            }
                            if ((Attachdirectories[i].Contains("apl") ? true : Attachdirectories[i].Contains("apl")))
                            {
                                GetAPL_Files();
                            }
                            if ((Attachdirectories[i].Contains("learmexico") ? true : Attachdirectories[i].Contains("learmexico")))
                            {
                                GetLear_Files();
                            }
                            if ((Attachdirectories [i].Contains("stellantis") ? true : Attachdirectories [i].Contains("stellantis")))
                            {
                                GetStellantis_Files();
                            }
                            if ((Attachdirectories[i].Contains("stellantis-mts") ? true : Attachdirectories[i].Contains("stellantis-mts")))
                            {
                                GetStellantisMTS_Files();
                            }
                            if ((Attachdirectories [i].Contains("dhlsupply") ? true : Attachdirectories [i].Contains("dhlsupply")))
                            {
                                GetDHLAttachment_Files();
                            }
                            if ((Attachdirectories [i].Contains("schneider") ? true : Attachdirectories [i].Contains("schneider")))
                            {
                                Getschneider_Files();
                            }
                            if ((Attachdirectories [i].Contains("jbhunt") ? true : Attachdirectories [i].Contains("jbhunt")))
                            {
                                GetJBHunt_Files();
                            }
                            if ((Attachdirectories [i].Contains("truper") ? true : Attachdirectories [i].Contains("truper")))
                            {
                                GetTruper_Files();
                            }
                            if ((Attachdirectories [i].Contains("tremec") ? true : Attachdirectories [i].Contains("tremec")))
                            {
                                GetTremec_Files();
                            }
                            if ((Attachdirectories[i].Contains("transplace") ? true : Attachdirectories[i].Contains("transplace")))
                            {
                                GetXml_Files();
                            }
                        }
                        else
                        {
                            Console.WriteLine("Source path does not exist!");
                        }
                    }

                }

                

                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Leer Carpetas...", erro));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void GetXml_Files()
        {
            // Lectua del xml
            //string xml = File.ReadAllText(@"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test\830059015.xml");


            Regex regexObj = new Regex(@"[^\d]");
            Regex regexObjdec = new Regex(@"^[0-9.-]+$");
            string Municipiofmt = "000";
            string CPfmt = "00000";
            string xmlSucio;
            string xmlLimpio;

            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            var directories = @"\\10.1.1.30\e$\Attachments\sftr-transplace";
            bool direxists = System.IO.Directory.Exists(directories);

            try
            {
                if (direxists)
                {
                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xml") ? true : fileInfo.Name.Contains(".xml")))
                        { 
                        
                            xmlSucio = LimpiaXMLDocs(fileInfo.DirectoryName + "\\" + fileInfo.Name);
                            xmlLimpio = RemoveSpecialCharacters(xmlSucio);
                            SaveXmlCleanDoc(xmlLimpio, fileInfo.DirectoryName + "\\" + fileInfo.Name);

                            string xml = File.ReadAllText(fileInfo.DirectoryName + "\\" + fileInfo.Name);
                            cfdi comp = new cfdi();
                            comp = XmlDeserializeFromString<cfdi>(xml);
                            if (comp == null)
                                MessageBox.Show("Error El archivo XML no es correcto");
                            else
                            {
                                //MessageBox.Show($"Load: {comp.Load}");
                                //MessageBox.Show($"Nombre del Receptor {comp.Receptor.Receptor_Nombre}");
                                string mercancias = string.Empty;
                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (cfdiConcepto concepto in comp.Conceptos)
                                {

                                    EntintDHLFields ent = new EntintDHLFields();


                                    DataTable ReferenciaSat = new DataTable();
                                    DataTable ReferenciaSat2 = new DataTable();

                                    ent.CP = comp.Emisor.Emisor_codigoPostal.ToString();
                                    ent.CP2 = comp.Receptor.Receptor_codigoPostal.ToString();

                                    ent.CP = String.IsNullOrEmpty(ent.CP) ? "00000" : ent.CP;
                                    ent.CP2 = String.IsNullOrEmpty(ent.CP2) ? "00000" : ent.CP2;

                                    ent.CP = Convert.ToInt32(regexObj.Replace(ent.CP.Replace(" ", ""), "")).ToString(CPfmt);
                                    ent.CP2 = Convert.ToInt32(regexObj.Replace(ent.CP2.Replace(" ", ""), "")).ToString(CPfmt);


                                    ReferenciaSat = ObtenerReferenciaSat(ent.CP);
                                    ReferenciaSat2 = ObtenerReferenciaSat(ent.CP2);

                                    mercancias = mercancias + $"  {concepto.SAT_Descripcion} --  {concepto.SAT_ClaveProdServ}";

                                    ent.ReferenciaDelServicio = comp.Load.ToString();
                                    ent.RSdelRemitente = comp.Load.ToString();
                                    ent.RFCdelRemitente = comp.Emisor.Emisor_Rfc;
                                    ent.Supplier = comp.Emisor.Emisor_Rfc;
                                    ent.Calle = comp.Emisor.Emisor_Direccion;
                                    ent.Municipio = ReferenciaSat.Rows[0][2].ToString();
                                    ent.Estado = ReferenciaSat.Rows[0][1].ToString();
                                    ent.Pais = "MEX";
                                    ent.RSdelDestinatario = comp.Receptor.Receptor_Rfc;
                                    ent.RFCDestinatario = comp.Receptor.Receptor_Rfc;
                                    ent.Calle2 = comp.Receptor.Receptor_Direccion;
                                    ent.Municipio2 = ReferenciaSat2.Rows[0][2].ToString();
                                    ent.Estado2 = ReferenciaSat2.Rows[0][1].ToString();
                                    ent.Pais2 = "MEX";
                                    ent.PesoNeto = concepto.Peso.ToString();
                                    ent.NumeroTotalMercancias = concepto.Cantidad.ToString();
                                    ent.ClaveDelBienTransportado = concepto.SAT_ClaveProdServ.ToString();
                                    ent.ClaveUnidadDeMedida = concepto.UnidadMedidaPeso;
                                    ent.DescripcionDelBienTransportado = concepto.SAT_Descripcion;
                                    ent.MaterialPeligroso = "No";
                                    ent.ValorDeLaMercancia = "1";
                                    ent.TipoDeMoneda = comp.Moneda;

                                    if (ent.ReferenciaDelServicio != "")
                                        eDhlList.Add(ent);

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                        }
                    }
                }
                CortarDocumentosXML(directories);
            }
            catch (Exception)
            {

                throw;
            }

        }


        public static T XmlDeserializeFromString<T>(string objectData)
        {
            return (T)XmlDeserializeFromString(objectData, typeof(T));
        }

        public static object XmlDeserializeFromString(string objectData, Type type)
        {
            var serializer = new XmlSerializer(type);
            object result;

            using (TextReader reader = new StringReader(objectData))
            {
                try
                {
                    result = serializer.Deserialize(reader);
                }
                catch (Exception ex)
                {
                    // MessageBox.Show(ex.Message + "Adicional:" + ex.InnerException);
                    result = null;
                }
            }

            return result;
        }

        private void GetDHLAttachment_Files()
        {
            string _fName = "";
            string strFullNme = "";

            int tb_EmbarqueDHL = 0;
            int tb_Orden = 1;
            int tb_IDOrigen = 1;
            int tb_RFCRemitente2 = 2;
            int tb_Calle = 3;
            int tb_Municipio = 5;
            int tb_Estado = 7;
            int tb_Pais = 8;
            int tb_CodigoPostal = 9;
            int tb_IDDestino = 10;
            int tb_RFCDestinatario2 = 11;
            int tb_Calle2 = 12;
            int tb_Municipio3 = 14;
            int tb_Estado4 = 16;
            int tb_Pais5 = 17;
            int tb_CodigoPostal6 = 18;
            int tb_PesoNetoTotal = 19;
            int tb_NumTotalMercancias = 20;
            int tb_BienesTransp = 21;
            int tb_Descripción = 22;
            int tb_ClaveUnidad = 23;
            int tb_CveMaterialPeligroso = 24;
            int tb_Embalaje = 22;
            int tb_DescripEmbalaje = 23;
            int tb_ValorMercancia = 28;
            int tb_FraccionArancelaria = 30;
            int tb_UUIDComercioExt = 26;
            int tb_TotalKMRuta = 27;
            int InicioHeader = 7;
            int InicioTabla = 8;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();



            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            var directories = @"\\10.1.1.30\e$\Attachments\sftr-dhlsupply\";
            bool direxists = System.IO.Directory.Exists(directories);

            try
            {
                // Setup session options
                

                if (direxists)
                {
                    //DataTable testOb = new DataTable();
                    //testOb = ObtenerOrigenCarga(28358);
                    //var directories = Directory.GetDirectories(@"\\10.1.1.30\FTProot\sftr-dhlsupply");
                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            //if ("NO" == EstaOnoBaneado(fileInfo.FullName))
                            //{

                            //}

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;

                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla

                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_Municipio].ToString() != null) && (e.ItemArray[tb_Estado].ToString() != null)
                                        && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                && (e.ItemArray[tb_Pais].ToString() != null) && (e.ItemArray[tb_CodigoPostal].ToString() != null) && (e.ItemArray[tb_IDDestino].ToString() != null)
                                                && (e.ItemArray[tb_Municipio3].ToString() != null) && (e.ItemArray[tb_Estado4].ToString() != null) && (e.ItemArray[tb_Pais5].ToString() != null)
                                                && (e.ItemArray[tb_CodigoPostal6].ToString() != null) && (e.ItemArray[tb_PesoNetoTotal].ToString() != null) && (e.ItemArray[tb_NumTotalMercancias].ToString() != null)
                                                && (e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_ClaveUnidad].ToString() != null) && (e.ItemArray[tb_Descripción].ToString() != null)
                                                && (excelDT.Columns["IDORIGEN"].Ordinal == tb_IDOrigen) && (excelDT.Columns["DESCRIPCION"].Ordinal == tb_Descripción))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            if (e.ItemArray[tb_EmbarqueDHL].ToString().Contains('.'))
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString().Split('.')[1]; }
                                            else
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString(); }
                                            ent.RSdelRemitente = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = e.ItemArray[tb_Municipio].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString().Trim(new Char[] { '\'' });
                                            ent.RSdelDestinatario = e.ItemArray[tb_IDDestino].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString().Trim(new Char[] { '\'' });
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString().Trim(new Char[] { '\'' });
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString().Trim(new Char[] { '\'' });
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                                            break;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);
                }

                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);
                GetDHLAttachmentSecondTry_Files();
                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion DHL...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        private void GetDHLAttachmentSecondTry_Files()
        {
            string _fName = "";
            string strFullNme = "";

            int tb_EmbarqueDHL = 0;
            int tb_Orden = 1;
            int tb_IDOrigen = 1;
            int tb_RFCRemitente2 = 2;
            int tb_Calle = 3;
            int tb_Municipio = 5;
            int tb_Estado = 7;
            int tb_Pais = 8;
            int tb_CodigoPostal = 9;
            int tb_IDDestino = 10;
            int tb_RFCDestinatario2 = 11;
            int tb_Calle2 = 12;
            int tb_Municipio3 = 14;
            int tb_Estado4 = 16;
            int tb_Pais5 = 17;
            int tb_CodigoPostal6 = 18;
            int tb_PesoNetoTotal = 19;
            int tb_NumTotalMercancias = 20;
            int tb_BienesTransp = 21;
            int tb_Descripción = 22;
            int tb_ClaveUnidad = 23;
            int tb_CveMaterialPeligroso = 24;
            int tb_Embalaje = 22;
            int tb_DescripEmbalaje = 23;
            int tb_ValorMercancia = 28;
            int tb_FraccionArancelaria = 30;
            int tb_UUIDComercioExt = 26;
            int tb_TotalKMRuta = 27;
            int InicioHeader = 8;
            int InicioTabla = 9;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();



            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            var directories = @"\\10.1.1.30\e$\Attachments\sftr-dhlsupply\";
            bool direxists = System.IO.Directory.Exists(directories);

            try
            {
                // Setup session options


                if (direxists)
                {
                    //DataTable testOb = new DataTable();
                    //testOb = ObtenerOrigenCarga(28358);
                    //var directories = Directory.GetDirectories(@"\\10.1.1.30\FTProot\sftr-dhlsupply");
                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            //if ("NO" == EstaOnoBaneado(fileInfo.FullName))
                            //{

                            //}

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;

                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla

                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_Municipio].ToString() != null) && (e.ItemArray[tb_Estado].ToString() != null)
                                        && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                && (e.ItemArray[tb_Pais].ToString() != null) && (e.ItemArray[tb_CodigoPostal].ToString() != null) && (e.ItemArray[tb_IDDestino].ToString() != null)
                                                && (e.ItemArray[tb_Municipio3].ToString() != null) && (e.ItemArray[tb_Estado4].ToString() != null) && (e.ItemArray[tb_Pais5].ToString() != null)
                                                && (e.ItemArray[tb_CodigoPostal6].ToString() != null) && (e.ItemArray[tb_PesoNetoTotal].ToString() != null) && (e.ItemArray[tb_NumTotalMercancias].ToString() != null)
                                                && (e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_ClaveUnidad].ToString() != null) && (e.ItemArray[tb_Descripción].ToString() != null)
                                                && (excelDT.Columns["IDORIGEN"].Ordinal == tb_IDOrigen) && (excelDT.Columns["DESCRIPCION"].Ordinal == tb_Descripción))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            if (e.ItemArray[tb_EmbarqueDHL].ToString().Contains('.'))
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString().Split('.')[1]; }
                                            else
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString(); }
                                            ent.RSdelRemitente = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = e.ItemArray[tb_Municipio].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString().Trim(new Char[] { '\'' });
                                            ent.RSdelDestinatario = e.ItemArray[tb_IDDestino].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString().Trim(new Char[] { '\'' });
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString().Trim(new Char[] { '\'' });
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString().Trim(new Char[] { '\'' });
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                                            break;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);
                }

                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));

                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion DHL...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void GetDHL_Files()
        {

            string _fName = "";
            
            string strFullNme = "";

            int tb_EmbarqueDHL = 0;
            int tb_Orden = 1;
            int tb_IDOrigen = 1;
            int tb_RFCRemitente2 = 2;
            int tb_Calle = 3;
            int tb_Municipio = 5;
            int tb_Estado = 7;
            int tb_Pais = 8;
            int tb_CodigoPostal = 9;
            int tb_IDDestino = 10;
            int tb_RFCDestinatario2 = 11;
            int tb_Calle2 = 12;
            int tb_Municipio3 = 14;
            int tb_Estado4 = 16;
            int tb_Pais5 = 17;
            int tb_CodigoPostal6 = 18;
            int tb_PesoNetoTotal = 19;
            int tb_NumTotalMercancias = 20;
            int tb_BienesTransp = 21;
            int tb_Descripción = 22;
            int tb_ClaveUnidad = 23;
            int tb_CveMaterialPeligroso = 24;
            int tb_Embalaje = 22;
            int tb_DescripEmbalaje = 23;
            int tb_ValorMercancia = 28;
            int tb_FraccionArancelaria = 30;
            int tb_UUIDComercioExt = 26;
            int tb_TotalKMRuta = 27;
            int InicioHeader = 7;
            int InicioTabla = 8;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();


            
            var directories = @"\\10.1.1.30\FTProot\sftr-dhlsupply";
            bool direxists = System.IO.Directory.Exists(directories);

            try
            {
                // Setup session options
                

                if (direxists)
                {

                    //var directories = Directory.GetDirectories(@"\\10.1.1.30\FTProot\sftr-dhlsupply");
                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;

                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla
                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_Municipio].ToString() != null) && (e.ItemArray[tb_Estado].ToString() != null)
                                            && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                    && (e.ItemArray[tb_Pais].ToString() != null) && (e.ItemArray[tb_CodigoPostal].ToString() != null) && (e.ItemArray[tb_IDDestino].ToString() != null)
                                                    && (e.ItemArray[tb_Municipio3].ToString() != null) && (e.ItemArray[tb_Estado4].ToString() != null) && (e.ItemArray[tb_Pais5].ToString() != null)
                                                    && (e.ItemArray[tb_CodigoPostal6].ToString() != null) && (e.ItemArray[tb_PesoNetoTotal].ToString() != null) && (e.ItemArray[tb_NumTotalMercancias].ToString() != null)
                                                    && (e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_ClaveUnidad].ToString() != null) && (e.ItemArray[tb_Descripción].ToString() != null)
                                                    && (excelDT.Columns["IDORIGEN"].Ordinal == tb_IDOrigen) && (excelDT.Columns["DESCRIPCION"].Ordinal == tb_Descripción))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            if (e.ItemArray[tb_EmbarqueDHL].ToString().Contains('.'))
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString().Split('.')[1]; }
                                            else
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString(); }
                                            ent.RSdelRemitente = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = e.ItemArray[tb_Municipio].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString().Trim(new Char[] { '\'' });
                                            ent.RSdelDestinatario = e.ItemArray[tb_IDDestino].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString().Trim(new Char[] { '\'' });
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString().Trim(new Char[] { '\'' });
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString().Trim(new Char[] { '\'' });
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                                            break;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);
                }

                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion DHL F...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        private void GetStellantisMTS_Files()
        {

            string _fName = "";
            
            string strFullNme = "";

            int  tb_UniqueRecordIdentifier = 0;
            int  tb_SenderID = 1;
            int  tb_ReciverID = 2;
            int  tb_MerchandiseOwnerTaxID = 3;
            int  tb_MerchandiseOwner = 4;
            int  tb_MerchandiseOwnerAddressLine1 = 5;
            int  tb_MerchandiseOwnerAddressLine2 = 6;
            int  tb_MerchandiseOwnerCity = 7;
            int  tb_MerchandiseOwnerState = 8;
            int  tb_MerchandiseOwnerZip = 9;
            int  tb_MerchandiseMunicipalySATreference = 10;
            int  tb_MerchandiseOwnerCntry = 11;
            int  tb_ShipTo = 12;
            int  tb_ShipToName = 13;
            int  tb_ShipToAddressLine = 14;
            int  tb_ShipToCity = 15;
            int  tb_ShipToState = 16;
            int  tb_ShipToZip = 17;
            int  tb_ShipToMunicipalySATreference = 18;
            int  tb_ShipToCntry = 19;
            int  tb_ShipFrom = 20;
            int  tb_SupplierTaxIdentifier = 21;
            int  tb_SupplierName = 22;
            int  tb_SupplierAddressLine = 23;
            int  tb_SupplierCity = 24;
            int  tb_SupplierState = 25;
            int  tb_SupplierZip = 26;
            int  tb_SupplierMuniciplalitySATreference = 27;
            int  tb_SupplierCntry = 28;
            int  tb_PartOrContainerID = 29;
            int  tb_PartDescription = 30;
            int  tb_PartSATCode = 31;
            int  tb_PartSATDescription = 32;
            int  tb_ShippedQuantity = 33;
            int  tb_UnitOfMeasureShipped = 34;
            int  tb_UnitOfMeasureSATCode = 35;
            int  tb_UnitOfMeasureSATdescription = 36;
            int  tb_HazmatFlag = 37;
            int  tb_HazmatSATCode = 38;
            int  tb_HazmatSATDescription = 39;
            int  tb_ContainerIdentifier = 40;
            int  tb_I_SAT_CNTNR = 41;
            int  tb_ContainerSATDescription = 42;
            int  tb_ContainerQty = 43;
            int  tb_ContainerTareWeight = 44;
            int  tb_NetShipmentWeight = 45;
            int  tb_GrossShipmentWeight = 46;
            int  tb_UnitOfMeasureWeight = 47;
            int  tb_HTSCode = 48;
            int  tb_HTSCountryCode = 49;
            int  tb_CurrencyCode = 50;
            int  tb_SupplierCode = 51;
            int  tb_FinalDestinationCode = 52;
            int  tb_ShipmentIdentifier = 53;
            int  tb_SupplierPackingSlip = 54;
            int  tb_SupplierBillOfLoading = 55;
            int  tb_FreightConsolidationBillOfLandingNumber = 56;
            int  tb_ConsolidationShipmentIdentifier = 57;
            int  tb_PoolpointShipfrom = 58;
            int  tb_PoolPointShipto = 59;
            int  tb_ShipmentDate = 60;
            int  tb_ShipmentTime = 61;
            int  tb_ShipmentTimestamp = 62;
            int  tb_CarrierSCAC = 63;
            int  tb_ConveyanceIdentifier = 64;
            int  tb_OwnerSCAC = 65;
            int  tb_TransportationMode = 66;
            int  tb_PartCountInContainer = 67;
            int  tb_AETCNumer = 68;
            int  tb_LotNumber = 69;
            int  tb_ChampsTransactionCode = 70;
            int  tb_ChampsPurposeCode = 71;
            int  tb_ASNStatus = 72;
            int  tb_ShipmentIdentifierCount = 73;
            int  tb_ASNCount = 74;
            int  tb_MasterBilOfLadinng = 75;
            int  tb_UnitEstimatedCost = 76;
            int  tb_FillerForFutureUser = 77;
            int  tb_TotalWeightContainer = 46;
            int  tb_TotalWeightContainerUnit = 47;
            int  tb_GrossShipmentWeightUnit = 47;
            int InicioHeader = 1;
            int InicioTabla = 2;


            int[] ordenColm = new int[] { tb_UniqueRecordIdentifier, tb_SenderID, tb_ReciverID, tb_MerchandiseOwnerTaxID, tb_MerchandiseOwner, tb_MerchandiseOwnerAddressLine1, tb_MerchandiseOwnerAddressLine2, tb_MerchandiseOwnerCity, tb_MerchandiseOwnerState, tb_MerchandiseOwnerZip, tb_MerchandiseMunicipalySATreference, tb_MerchandiseOwnerCntry, tb_ShipTo, tb_ShipToName, tb_ShipToAddressLine, tb_ShipToCity, tb_ShipToState, tb_ShipToZip, tb_ShipToMunicipalySATreference, tb_ShipToCntry, tb_ShipFrom, tb_SupplierTaxIdentifier, tb_SupplierName, tb_SupplierAddressLine, tb_SupplierCity, tb_SupplierState, tb_SupplierZip, tb_SupplierMuniciplalitySATreference, tb_SupplierCntry, tb_PartOrContainerID, tb_PartDescription, tb_PartSATCode, tb_PartSATDescription, tb_ShippedQuantity, tb_UnitOfMeasureShipped, tb_UnitOfMeasureSATCode, tb_UnitOfMeasureSATdescription, tb_HazmatFlag, tb_HazmatSATCode, tb_HazmatSATDescription, tb_ContainerIdentifier, tb_I_SAT_CNTNR, tb_ContainerSATDescription, tb_ContainerQty, tb_ContainerTareWeight, tb_NetShipmentWeight, tb_GrossShipmentWeight, tb_UnitOfMeasureWeight, tb_HTSCode, tb_HTSCountryCode, tb_CurrencyCode, tb_SupplierCode, tb_FinalDestinationCode, tb_ShipmentIdentifier, tb_SupplierPackingSlip, tb_SupplierBillOfLoading, tb_FreightConsolidationBillOfLandingNumber, tb_ConsolidationShipmentIdentifier, tb_PoolpointShipfrom, tb_PoolPointShipto, tb_ShipmentDate, tb_ShipmentTime, tb_ShipmentTimestamp, tb_CarrierSCAC, tb_ConveyanceIdentifier, tb_OwnerSCAC, tb_TransportationMode, tb_PartCountInContainer, tb_AETCNumer, tb_LotNumber, tb_ChampsTransactionCode, tb_ChampsPurposeCode, tb_ASNStatus, tb_ShipmentIdentifierCount, tb_ASNCount, tb_MasterBilOfLadinng, tb_UnitEstimatedCost, tb_FillerForFutureUser };
            DataTable errorTable = new DataTable();



            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            var directories = @"\\10.1.1.30\e$\Attachments\sftr-stellantis-mts";
            bool direxists = System.IO.Directory.Exists(directories);

            try
            {
                // Setup session options
                

                if (direxists)
                {

                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;

                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla

                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }

                                List<EntityStellantis> eDhlList = new List<EntityStellantis>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if (((e.ItemArray[tb_UniqueRecordIdentifier].ToString() != null) && ((e.ItemArray[tb_UniqueRecordIdentifier].ToString() != "")
                                            && (e.ItemArray[tb_MerchandiseOwnerTaxID].ToString() != null)) && (e.ItemArray[tb_MerchandiseOwnerTaxID].ToString() != ""))
                                            && ((e.ItemArray[tb_MerchandiseOwnerAddressLine1].ToString() != null) && (e.ItemArray[tb_MerchandiseOwnerAddressLine1].ToString() != ""))
                                                && (excelDT.Columns["Unique Record Identifier"].Ordinal == tb_UniqueRecordIdentifier) && (excelDT.Columns["Receiver ID"].Ordinal == tb_ReciverID))
                                        {

                                            EntityStellantis ent = new EntityStellantis();

                                            ent.UniqueRecordIdentifier = e.ItemArray[tb_UniqueRecordIdentifier].ToString();
                                            ent.SenderID = e.ItemArray[tb_SenderID].ToString();
                                            ent.ReciverID = e.ItemArray[tb_ReciverID].ToString();
                                            ent.MerchandiseOwnerTaxID = e.ItemArray[tb_MerchandiseOwnerTaxID].ToString();
                                            ent.MerchandiseOwner = e.ItemArray[tb_MerchandiseOwner].ToString();
                                            ent.MerchandiseOwnerAddressLine1 = e.ItemArray[tb_MerchandiseOwnerAddressLine1].ToString();
                                            ent.MerchandiseOwnerAddressLine2 = e.ItemArray[tb_MerchandiseOwnerAddressLine2].ToString();
                                            ent.MerchandiseOwnerCity = e.ItemArray[tb_MerchandiseOwnerCity].ToString();
                                            ent.MerchandiseOwnerState = e.ItemArray[tb_MerchandiseOwnerState].ToString();
                                            ent.MerchandiseOwnerZip = e.ItemArray[tb_MerchandiseOwnerZip].ToString();
                                            ent.MerchandiseMunicipalySATreference = e.ItemArray[tb_MerchandiseMunicipalySATreference].ToString();
                                            ent.MerchandiseOwnerCntry = e.ItemArray[tb_MerchandiseOwnerCntry].ToString();
                                            ent.ShipTo = e.ItemArray[tb_ShipTo].ToString();
                                            ent.ShipToName = e.ItemArray[tb_ShipToName].ToString();
                                            ent.ShipToAddressLine = e.ItemArray[tb_ShipToAddressLine].ToString();
                                            ent.ShipToCity = e.ItemArray[tb_ShipToCity].ToString();
                                            ent.ShipToState = e.ItemArray[tb_ShipToState].ToString();
                                            ent.ShipToZip = e.ItemArray[tb_ShipToZip].ToString();
                                            ent.ShipToMunicipalySATreference = e.ItemArray[tb_ShipToMunicipalySATreference].ToString();
                                            ent.ShipToCntry = e.ItemArray[tb_ShipToCntry].ToString();
                                            ent.ShipFrom = e.ItemArray[tb_ShipFrom].ToString();
                                            ent.SupplierTaxIdentifier = e.ItemArray[tb_SupplierTaxIdentifier].ToString();
                                            ent.SupplierName = e.ItemArray[tb_SupplierName].ToString();
                                            ent.SupplierAddressLine = e.ItemArray[tb_SupplierAddressLine].ToString();
                                            ent.SupplierCity = e.ItemArray[tb_SupplierCity].ToString();
                                            ent.SupplierState = e.ItemArray[tb_SupplierState].ToString();
                                            ent.SupplierZip = e.ItemArray[tb_SupplierZip].ToString();
                                            ent.SupplierMuniciplalitySATreference = e.ItemArray[tb_SupplierMuniciplalitySATreference].ToString();
                                            ent.SupplierCntry = e.ItemArray[tb_SupplierCntry].ToString();
                                            ent.PartOrContainerID = e.ItemArray[tb_PartOrContainerID].ToString();
                                            ent.PartDescription = e.ItemArray[tb_PartDescription].ToString();
                                            ent.PartSATCode = e.ItemArray[tb_PartSATCode].ToString();
                                            ent.PartSATDescription = e.ItemArray[tb_PartSATDescription].ToString();
                                            ent.ShippedQuantity = e.ItemArray[tb_ShippedQuantity].ToString();
                                            ent.UnitOfMeasureShipped = e.ItemArray[tb_UnitOfMeasureShipped].ToString();
                                            ent.UnitOfMeasureSATCode = e.ItemArray[tb_UnitOfMeasureSATCode].ToString();
                                            ent.UnitOfMeasureSATdescription = e.ItemArray[tb_UnitOfMeasureSATdescription].ToString();
                                            ent.HazmatFlag = e.ItemArray[tb_HazmatFlag].ToString();
                                            ent.HazmatSATCode = e.ItemArray[tb_HazmatSATCode].ToString();
                                            ent.HazmatSATDescription = e.ItemArray[tb_HazmatSATDescription].ToString();
                                            ent.ContainerIdentifier = e.ItemArray[tb_ContainerIdentifier].ToString();
                                            ent.I_SAT_CNTNR = e.ItemArray[tb_I_SAT_CNTNR].ToString();
                                            ent.ContainerSATDescription = e.ItemArray[tb_ContainerSATDescription].ToString();
                                            ent.ContainerQty = e.ItemArray[tb_ContainerQty].ToString();
                                            ent.ContainerTareWeight = e.ItemArray[tb_ContainerTareWeight].ToString();
                                            ent.NetShipmentWeight = e.ItemArray[tb_NetShipmentWeight].ToString();
                                            ent.GrossShipmentWeight = e.ItemArray[tb_GrossShipmentWeight].ToString();
                                            ent.UnitOfMeasureWeight = e.ItemArray[tb_UnitOfMeasureWeight].ToString();
                                            ent.HTSCode = e.ItemArray[tb_HTSCode].ToString();
                                            ent.HTSCountryCode = e.ItemArray[tb_HTSCountryCode].ToString();
                                            ent.CurrencyCode = e.ItemArray[tb_CurrencyCode].ToString();
                                            ent.SupplierCode = e.ItemArray[tb_SupplierCode].ToString();
                                            ent.FinalDestinationCode = e.ItemArray[tb_FinalDestinationCode].ToString();
                                            ent.ShipmentIdentifier = e.ItemArray[tb_ShipmentIdentifier].ToString();
                                            ent.SupplierPackingSlip = e.ItemArray[tb_SupplierPackingSlip].ToString();
                                            ent.SupplierBillOfLoading = e.ItemArray[tb_SupplierBillOfLoading].ToString();
                                            ent.FreightConsolidationBillOfLandingNumber = e.ItemArray[tb_FreightConsolidationBillOfLandingNumber].ToString();
                                            ent.ConsolidationShipmentIdentifier = e.ItemArray[tb_ConsolidationShipmentIdentifier].ToString();
                                            ent.PoolpointShipfrom = e.ItemArray[tb_PoolpointShipfrom].ToString();
                                            ent.PoolPointShipto = e.ItemArray[tb_PoolPointShipto].ToString();
                                            ent.ShipmentDate = e.ItemArray[tb_ShipmentDate].ToString();
                                            ent.ShipmentTime = e.ItemArray[tb_ShipmentTime].ToString();
                                            ent.ShipmentTimestamp = e.ItemArray[tb_ShipmentTimestamp].ToString();
                                            ent.CarrierSCAC = e.ItemArray[tb_CarrierSCAC].ToString();
                                            ent.ConveyanceIdentifier = e.ItemArray[tb_ConveyanceIdentifier].ToString();
                                            ent.OwnerSCAC = e.ItemArray[tb_OwnerSCAC].ToString();
                                            ent.TransportationMode = e.ItemArray[tb_TransportationMode].ToString();
                                            ent.PartCountInContainer = e.ItemArray[tb_PartCountInContainer].ToString();
                                            ent.AETCNumer = e.ItemArray[tb_AETCNumer].ToString();
                                            ent.LotNumber = e.ItemArray[tb_LotNumber].ToString();
                                            ent.ChampsTransactionCode = e.ItemArray[tb_ChampsTransactionCode].ToString();
                                            ent.ChampsPurposeCode = e.ItemArray[tb_ChampsPurposeCode].ToString();
                                            ent.ASNStatus = e.ItemArray[tb_ASNStatus].ToString();
                                            ent.ShipmentIdentifierCount = e.ItemArray[tb_ShipmentIdentifierCount].ToString();
                                            ent.ASNCount = e.ItemArray[tb_ASNCount].ToString();
                                            ent.MasterBilOfLadinng = e.ItemArray[tb_MasterBilOfLadinng].ToString();
                                            ent.UnitEstimatedCost = "";
                                            ent.FillerForFutureUser = "";
                                            ent.TotalWeightContainer = e.ItemArray[tb_TotalWeightContainer].ToString();
                                            ent.TotalWeightContainerUnit = e.ItemArray[tb_TotalWeightContainerUnit].ToString();
                                            ent.GrossShipmentWeightUnit = e.ItemArray[tb_GrossShipmentWeightUnit].ToString();

                                            if (ent.UniqueRecordIdentifier != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                            break;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }
                                }
                                if (eDhlList.Any())
                                {
                                    List<EntityStellantis> StellantisL = new List<EntityStellantis>();
                                    StellantisL = VerificaMercanciasStellantis(eDhlList);
                                    bool isEmpt = !StellantisL.Any();
                                    if (isEmpt){}
                                    else
                                    {
                                        InsertaMercanciaClientesStellantisList(StellantisL, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);

                }
                //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion StellantisMTS...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }



        private void GetGeneric_Files()
        {

            string _fName = "";
            
            string strFullNme = "";

            int tb_EmbarqueDHL = 0;
            int tb_Orden = 1;
            int tb_RSremitente = 1;
            int tb_IDOrigen = 1;
            int tb_RFCRemitente2 = 2;
            int tb_Calle = 3;
            int tb_Municipio = 4;
            int tb_Estado = 5;
            int tb_Pais = 6;
            int tb_CodigoPostal = 7;
            int tb_IDDestino = 10;
            int tb_RSDest = 8;
            int tb_RFCDestinatario2 = 9;
            int tb_Calle2 = 10;
            int tb_Municipio3 = 11;
            int tb_Estado4 = 12;
            int tb_Pais5 = 13;
            int tb_CodigoPostal6 = 14;
            int tb_PesoNetoTotal = 15;
            int tb_NumTotalMercancias = 16;
            int tb_BienesTransp = 17;
            int tb_Descripción = 18;
            int tb_ClaveUnidad = 19;
            int tb_CveMaterialPeligroso = 20;
            int tb_Embalaje = 22;
            int tb_DescripEmbalaje = 23;
            int tb_ValorMercancia = 21;
            int tb_FraccionArancelaria = 22;
            int tb_UUIDComercioExt = 26;
            int tb_TotalKMRuta = 27;
            int InicioHeader = 1;
            int InicioTabla = 2;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();


            
            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            var directories = @"\\10.1.1.30\e$\Attachments\sftr-generico";
            bool direxists = System.IO.Directory.Exists(directories);

            try
            {
                // Setup session options
                

                if (direxists)
                {

                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;
                            
                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;

                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla

                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if (((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && ((e.ItemArray[tb_EmbarqueDHL].ToString() != "")
                                            &&(e.ItemArray[tb_RFCDestinatario2].ToString() != null)) && (e.ItemArray[tb_RFCDestinatario2].ToString() != ""))
                                            && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                && (excelDT.Columns["Calle"].Ordinal == tb_Calle) && (excelDT.Columns["Estado"].Ordinal == tb_Estado))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            if (e.ItemArray[tb_EmbarqueDHL].ToString().Contains('.'))
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString().Split('.')[1]; }
                                            else
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString(); }
                                            ent.RSdelRemitente = e.ItemArray[tb_RSremitente].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = e.ItemArray[tb_Municipio].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString().Trim(new Char[] { '\'' });
                                            ent.RSdelDestinatario = e.ItemArray[tb_RSDest].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString().Trim(new Char[] { '\'' });
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString().Trim(new Char[] { '\'' });
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString().Trim(new Char[] { '\'' });
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                            break;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }
                                }
                                if (eDhlList.Any()) {
                                    {
                                        List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                        ClientList = VerificaMercanciasClientes(eDhlList);
                                        bool isEmpt = !ClientList.Any();
                                        if (isEmpt) { }
                                        else
                                        {
                                            InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                        }
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);

                }
                //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion GenericF...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        private void GetAPL_Files()
        {

            string _fName = "";
            
            string strFullNme = "";

            int tb_EmbarqueDHL = 0;
            int tb_Orden = 1;
            int tb_RSremitente = 1;
            int tb_IDOrigen = 1;
            int tb_RFCRemitente2 = 2;
            int tb_Calle = 3;
            int tb_Municipio = 4;
            int tb_Estado = 5;
            int tb_Pais = 6;
            int tb_CodigoPostal = 7;
            int tb_IDDestino = 10;
            int tb_RSDest = 8;
            int tb_RFCDestinatario2 = 9;
            int tb_Calle2 = 10;
            int tb_Municipio3 = 11;
            int tb_Estado4 = 12;
            int tb_Pais5 = 13;
            int tb_CodigoPostal6 = 14;
            int tb_PesoNetoTotal = 15;
            int tb_NumTotalMercancias = 16;
            int tb_BienesTransp = 17;
            int tb_Descripción = 18;
            int tb_ClaveUnidad = 19;
            int tb_CveMaterialPeligroso = 20;
            int tb_Embalaje = 22;
            int tb_DescripEmbalaje = 23;
            int tb_ValorMercancia = 21;
            int tb_FraccionArancelaria = 22;
            int tb_UUIDComercioExt = 26;
            int tb_TotalKMRuta = 27;
            int InicioHeader = 1;
            int InicioTabla = 2;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();


            
            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            var directories = @"\\10.1.1.30\e$\Attachments\sftr-apl";
            bool direxists = System.IO.Directory.Exists(directories);

            try
            {
                // Setup session options
                

                if (direxists)
                {

                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;

                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla

                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if (((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && ((e.ItemArray[tb_EmbarqueDHL].ToString() != "")
                                            && (e.ItemArray[tb_RFCDestinatario2].ToString() != null)) && (e.ItemArray[tb_RFCDestinatario2].ToString() != ""))
                                            && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                && (excelDT.Columns["Calle"].Ordinal == tb_Calle) && (excelDT.Columns["Estado"].Ordinal == tb_Estado))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            if (e.ItemArray[tb_EmbarqueDHL].ToString().Contains('.'))
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString().Split('.')[1]; }
                                            else
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString(); }
                                            ent.RSdelRemitente = e.ItemArray[tb_RSremitente].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = e.ItemArray[tb_Municipio].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString().Trim(new Char[] { '\'' });
                                            ent.RSdelDestinatario = e.ItemArray[tb_RSDest].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString().Trim(new Char[] { '\'' });
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString().Trim(new Char[] { '\'' });
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString().Trim(new Char[] { '\'' });
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                            break;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }
                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);

                }
                //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion APLFiles...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }



        private void Getschneider_Files()
        {

            string _fName = "";
            
            string strFullNme = "";
            string SinValor = "000";

            int tb_EmbarqueDHL = 0;
            int tb_Orden = 2;
            int tb_IDOrigen = 12;
            //SinValor
            int tb_RFCRemitente2 = 12;
            int tb_Calle = 3;
            int tb_Municipio = 8;
            //SinValor
            int tb_Estado = 4;
            int tb_Pais = 5;
            int tb_CodigoPostal = 6;
            int tb_IDDestino = 13;
            int tb_RFCDestinatario2 = 13;
            int tb_Calle2 = 8;
            int tb_Municipio3 = 8;
            //SinValor
            int tb_Estado4 = 9;
            //SinValor
            int tb_Pais5 = 10;
            int tb_CodigoPostal6 = 11;
            int tb_PesoNetoTotal = 19;
            int tb_NumTotalMercancias = 14;
            int tb_BienesTransp = 15;
            int tb_Descripción = 16;
            int tb_ClaveUnidad = 18;
            int tb_CveMaterialPeligroso = 21;
            int tb_Embalaje = 22;
            int tb_DescripEmbalaje = 24;
            int tb_ValorMercancia = 24;     
            //SinValor
            int tb_FraccionArancelaria = 25;
            //SinValor
            int tb_UUIDComercioExt = 27;
            int tb_TotalKMRuta = 28;
            int InicioHeader = 5;
            int InicioTabla = 6;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();



            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            var directories = @"\\10.1.1.30\e$\Attachments\sftr-schneider";
            //var directories = @"\\10.1.1.30\e$\Attachments\sftr-upds";
            bool direxists = System.IO.Directory.Exists(directories);


            
            try
            {
                // Setup session options
                

                if (direxists)
                {

                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 3, InicioHeader, InicioTabla);
                                errorTable = excelDT;

                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla
                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if (((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_RFCDestinatario2].ToString() != ""))
                                            && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                    && (excelDT.Columns["Calle"].Ordinal == tb_Calle))
                                        {
                                            EntintDHLFields ent = new EntintDHLFields();

                                            ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString();
                                            ent.RSdelRemitente = SinValor;
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = SinValor;
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString();
                                            ent.RSdelDestinatario = SinValor;
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = SinValor;
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString();
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString();
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString();
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if(ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                                            Thread.Sleep(5300);
                                            break;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);
                }
                //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");

                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);
                Thread.Sleep(5300);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Schneider...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void GetJBHunt_Files()
        {

            string _fName = "";
            
            string strFullNme = "";
            string sinValor = "000";


            int tb_RSRemint = 14;
            int tb_RSDest = 21;
            int tb_EmbarqueDHL = 3;
            int tb_Orden = 19;
            int tb_IDOrigen = 12;
            int tb_RFCRemitente2 = 13;
            int tb_Calle = 15;
            int tb_Municipio = 12;
            int tb_Estado = 18;
            int tb_Pais = 17;
            int tb_CodigoPostal = 16;
            int tb_IDDestino = 19;
            int tb_RFCDestinatario2 = 20;
            int tb_Calle2 = 22;
            int tb_Municipio3 = 19;
            int tb_Estado4 = 19;
            int tb_Pais5 = 19;
            int tb_CodigoPostal6 = 20;
            int tb_PesoNetoTotal = 34;
            int tb_NumTotalMercancias = 23;
            int tb_BienesTransp = 31;
            int tb_Descripción = 24;
            int tb_ClaveUnidad = 32;
            int tb_CveMaterialPeligroso = 27;
            int tb_Embalaje = 24;
            int tb_DescripEmbalaje = 24;
            int tb_ValorMercancia = 23;
            //SinValor
            int tb_FraccionArancelaria = 26;
            //SinValor
            int tb_UUIDComercioExt = 27;
            int tb_TotalKMRuta = 28;
            int InicioHeader = 4;
            int InicioTabla = 6;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();

            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";

            var directories = @"\\10.1.1.30\e$\Attachments\sftr-jbhunt";
            bool direxists = System.IO.Directory.Exists(directories);


            
            try
            {
                // Setup session options
                

                //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
                if (direxists)
                {

                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;
                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla
                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }
                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if (((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_RFCDestinatario2].ToString() != ""))
                                            && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                    && (excelDT.Columns["IDOrigen"].Ordinal == tb_IDOrigen) && (excelDT.Columns["RFCRemitente"].Ordinal == tb_RFCRemitente2) && (excelDT.Columns["MaterialPeligroso"].Ordinal == tb_CveMaterialPeligroso))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString();
                                            ent.RSdelRemitente = e.ItemArray[tb_RSRemint].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = sinValor;
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString();
                                            ent.RSdelDestinatario = e.ItemArray[tb_RSDest].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = sinValor;
                                            ent.Estado2 = sinValor;
                                            ent.Pais2 = sinValor;
                                            ent.CP2 = sinValor;
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString();
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                                            break;
                                        }

                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);
                }

                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);
                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion JBHunt...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void GetTruper_Files()
        {

            string _fName = "";
            
            string strFullNme = "";


            int tb_EmbarqueDHL = 0;
            int tb_Orden = 2;
            int tb_IDOrigen = 0;
            int tb_RFCRemitente2 = 2;
            //SinValor
            int tb_Calle = 3;
            int tb_Municipio = 4;
            //SinValor
            int tb_Estado = 5;
            int tb_Pais = 6;
            int tb_CodigoPostal = 7;
            int tb_IDDestino = 0;
            int tb_RFCDestinatario2 = 9;
            //SinValor
            int tb_Calle2 = 10;
            int tb_Municipio3 = 11;
            int tb_Estado4 = 12;
            int tb_Pais5 = 13;
            int tb_CodigoPostal6 = 14;
            int tb_PesoNetoTotal = 15;
            int tb_NumTotalMercancias = 16;
            int tb_BienesTransp = 17;
            int tb_Descripción = 18;
            int tb_ClaveUnidad = 19;
            int tb_CveMaterialPeligroso = 20;
            int tb_Embalaje = 23;
            int tb_DescripEmbalaje = 24;
            int tb_ValorMercancia = 21;
            int tb_FraccionArancelaria = 22;
            int tb_UUIDComercioExt = 27;
            int tb_TotalKMRuta = 28;
            int InicioHeader = 1;
            int InicioTabla = 3;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();

            var directories = @"\\10.1.1.30\e$\Attachments\sftr-truper";
            bool direxists = System.IO.Directory.Exists(directories);


            
            try
            {
                // Setup session options
                
                if (direxists)
                {

                    //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;

                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }
                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_EmbarqueDHL].ToString() != "")
                                            && (e.ItemArray[tb_Municipio].ToString() != null) && (e.ItemArray[tb_Municipio].ToString() != "")
                                            && (e.ItemArray[tb_Municipio3].ToString() != null) && (e.ItemArray[tb_Municipio3].ToString() != "")
                                            && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                    && (excelDT.Columns["C.P."].Ordinal == tb_CodigoPostal) && (excelDT.Columns["Estado"].Ordinal == tb_Estado))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString().Replace(' ', '-');
                                            ent.RSdelRemitente = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = e.ItemArray[tb_Municipio].ToString();
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString();
                                            ent.RSdelDestinatario = e.ItemArray[tb_IDDestino].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString();
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString();
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString();
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString();
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                                            break;
                                        }

                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);

                }
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Truper...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void GetAndroid_Files()
        {

            string _fName = "";
            
            string strFullNme = "";


            int tb_EmbarqueDHL = 0;
            int tb_Orden = 2;
            int tb_IDOrigen = 0;
            int tb_RFCRemitente2 = 2;
            //SinValor
            int tb_Calle = 3;
            int tb_Municipio = 4;
            //SinValor
            int tb_Estado = 5;
            int tb_Pais = 6;
            int tb_CodigoPostal = 7;
            int tb_IDDestino = 0;
            int tb_RFCDestinatario2 = 9;
            //SinValor
            int tb_Calle2 = 10;
            int tb_Municipio3 = 11;
            int tb_Estado4 = 12;
            int tb_Pais5 = 13;
            int tb_CodigoPostal6 = 14;
            int tb_PesoNetoTotal = 15;
            int tb_NumTotalMercancias = 16;
            int tb_BienesTransp = 17;
            int tb_Descripción = 18;
            int tb_ClaveUnidad = 19;
            int tb_CveMaterialPeligroso = 20;
            int tb_Embalaje = 23;
            int tb_DescripEmbalaje = 24;
            int tb_ValorMercancia = 21;
            int tb_FraccionArancelaria = 22;
            int tb_UUIDComercioExt = 27;
            int tb_TotalKMRuta = 28;
            int InicioHeader = 1;
            int InicioTabla = 2;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();


            var directories = @"\\10.1.1.30\e$\Attachments\sftr-android";
            bool direxists = System.IO.Directory.Exists(directories);


            
            try
            {
                // Setup session options
                

                if (direxists)
                {

                    //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;
                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }
                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_EmbarqueDHL].ToString() != "")
                                            && (e.ItemArray[tb_Municipio].ToString() != null) && (e.ItemArray[tb_Municipio].ToString() != "")
                                            && (e.ItemArray[tb_Municipio3].ToString() != null) && (e.ItemArray[tb_Municipio3].ToString() != "")
                                            && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                    && (excelDT.Columns["C.P"].Ordinal == tb_CodigoPostal) && (excelDT.Columns["Estado"].Ordinal == tb_Estado))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString();
                                            ent.RSdelRemitente = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = e.ItemArray[tb_Municipio].ToString();
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString();
                                            ent.RSdelDestinatario = e.ItemArray[tb_IDDestino].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString();
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString();
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString();
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString();
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                                            break;
                                        }

                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);

                }
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Android...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void GetCeltic_Files()
        {

            string _fName = "";
            
            string strFullNme = "";


            int tb_EmbarqueDHL = 0;
            int tb_Orden = 2;
            int tb_IDOrigen = 0;
            int tb_RFCRemitente2 = 2;
            //SinValor
            int tb_Calle = 3;
            int tb_Municipio = 4;
            //SinValor
            int tb_Estado = 5;
            int tb_Pais = 6;
            int tb_CodigoPostal = 7;
            int tb_IDDestino = 0;
            int tb_RFCDestinatario2 = 9;
            //SinValor
            int tb_Calle2 = 10;
            int tb_Municipio3 = 11;
            int tb_Estado4 = 12;
            int tb_Pais5 = 13;
            int tb_CodigoPostal6 = 14;
            int tb_PesoNetoTotal = 15;
            int tb_NumTotalMercancias = 16;
            int tb_BienesTransp = 17;
            int tb_Descripción = 18;
            int tb_ClaveUnidad = 19;
            int tb_CveMaterialPeligroso = 20;
            int tb_Embalaje = 23;
            int tb_DescripEmbalaje = 24;
            int tb_ValorMercancia = 21;
            int tb_FraccionArancelaria = 22;
            int tb_UUIDComercioExt = 27;
            int tb_TotalKMRuta = 28;
            int InicioHeader = 1;
            int InicioTabla = 3;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();



            var directories = @"\\10.1.1.30\e$\Attachments\sftr-celtic";
            bool direxists = System.IO.Directory.Exists(directories);


            
            try
            {
                // Setup session options
                

                if (direxists)
                {

                    //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;
                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }
                                //El primer numero determina cuando comienzala hoja, despues  el header el tercer numero es donde comienza la tabla
                                //1 = hoja, 6 = header, 8 = comienzo de la tabla

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_EmbarqueDHL].ToString() != "")
                                            && (e.ItemArray[tb_Municipio].ToString() != null) && (e.ItemArray[tb_Municipio].ToString() != "")
                                            && (e.ItemArray[tb_Municipio3].ToString() != null) && (e.ItemArray[tb_Municipio3].ToString() != "")
                                            && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                    && (excelDT.Columns["C.P."].Ordinal == tb_CodigoPostal) && (excelDT.Columns["Estado"].Ordinal == tb_Estado))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString();
                                            ent.RSdelRemitente = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = e.ItemArray[tb_Municipio].ToString();
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString();
                                            ent.RSdelDestinatario = e.ItemArray[tb_IDDestino].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString();
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString();
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString();
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString();
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                                            break;
                                        }

                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);

                }
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Celtic...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void GetAmazon_Files()
        {

            Regex regexObj = new Regex(@"[^\d]");
            string _fName = "";
            string _fNameInfo = "";

            string strFullNme = "";
            string CPfmt = "00000";
            string CPfmt1 = "";
            string CPfmt2 = "";


            int tb_EmbarqueDHL = 0;

            int tb_Calle1 = 27;
            int tb_Calle2 = 32;
            int tb_CP1 = 26;
            int tb_CP2 = 31;
            int tb_PesoNetoTotal = 4;
            int tb_NumTotalMercancias = 13;
            int tb_BienesTransp = 24;
            int tb_Descripción = 12;
            int tb_ClaveUnidad = 5;
            int tb_FraccionArancelaria = 21;
            int tb_UUIDComercioExt = 27;
            int tb_TotalKMRuta = 28;


            var directories = @"\\10.1.1.30\e$\Attachments\sftr-amazon";
            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            bool direxists = System.IO.Directory.Exists(directories);



            try
            {
                // Setup session options
                

                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();
                if (direxists)
                {

                    FileInfo[] filesChek = (new DirectoryInfo(directories)).GetFiles();

                    if(filesChek.Length != 0)
                    { 


                        FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                        for (int j = 0; j < (int)files.Length; j++)
                        {
                            //FileInfo fileInfoCheck = filesChek[0];
                            FileInfo fileInfo = files[j];
                            _fName = fileInfo.Name;
                            _fNameInfo = fileInfo.FullName;
                            if ((fileInfo.Name.Contains("VRID ") ? true : fileInfo.Name.Contains("VRID ")))
                            {
                                _fName = _fName.ToString().Split(' ')[1];
                            }
                            if ((fileInfo.Name.Contains("VRID") ? true : fileInfo.Name.Contains("VRID")))
                            {
                                _fName = _fName.ToString().Split('D')[1];
                            }
                            if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                            {
                                _fName = _fName.ToString().Split('.')[0];
                            }


                            if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")) && (fileInfo.Name.Contains(_fName) ? true : fileInfo.Name.Contains(_fName)))
                            {
                                if ((fileInfo.Name.Contains("-") ? true : fileInfo.Name.Contains("-")))
                                {
                                    _fName = _fName.ToString().Split('-')[0];
                                }
                                String filenameToUpload = fileInfo.Name;
                                strFullNme = fileInfo.FullName;

                                if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                                {
                                    DataTable excelDT = new DataTable();
                                    excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, 1, 2);

                                    //El primer numero determina cuando comienzala hoja, despues  el header el tercer numero es donde comienza la tabla
                                    //1 = hoja, 6 = header, 8 = comienzo de la tabla

                                

                                    foreach (DataRow e in excelDT.Rows)
                                    {
                                        if (excelDT.Rows.Count > 1)
                                        {
                                            if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_EmbarqueDHL].ToString() != "")
                                                    && (excelDT.Columns["Title"].Ordinal == tb_Descripción))
                                            {
                                                EntintDHLFields ent = new EntintDHLFields();

                                                ent.CP = "54944";
                                                ent.CP2 = "53319";


                                                ent.ReferenciaDelServicio = _fName;
                                                ent.RSdelRemitente = "SERVICIOS COMERCIALES AMAZON MEXICO S. DE R.L.";
                                                ent.RFCdelRemitente = "ANE140618P37";
                                                ent.Supplier = "ANE140618P37";
                                                ent.Calle = "Av. José López Portillo #92";
                                                ent.Municipio = "";
                                                ent.Estado = "MEX";
                                                ent.Pais = "MEX";
                                                ent.RSdelDestinatario = "SERVICIOS COMERCIALES AMAZON MEXICO S. DE R.L.";
                                                ent.RFCDestinatario = "ANE140618P37";
                                                ent.Calle2 = "San Jose de los Leones 5";
                                                ent.Municipio2 = "109";
                                                ent.Estado2 = "MEX";
                                                ent.Pais2 = "MEX";
                                                ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                                ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                                ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString().Trim(new Char[] { '\'' });
                                                ent.ClaveUnidadDeMedida = "KGM";
                                                ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString().Replace('|', ' ').Replace(',', ' ').Replace('°', ' ').Replace('-', ' ')
                                                    .Replace('\\', ' ').Replace('/', ' ').Replace('\'', ' ').Replace('’', ' ').Replace('"', ' ');
                                                ent.MaterialPeligroso = "No";
                                                ent.ValorDeLaMercancia = "1";
                                                ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                                if (ent.ReferenciaDelServicio != "")
                                                    eDhlList.Add(ent);

                                            }
                                            else
                                            {
                                                //Mandar email
                                                ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                                EnviarEmail(directories, fileInfo.Name, "Revisar VRID");
                                                break;
                                            }
                                        }
                                        else
                                        {

                                            if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_EmbarqueDHL].ToString() != "")
                                                    && (excelDT.Columns["Title"].Ordinal == tb_Descripción))
                                            {
                                                EntintDHLFields ent = new EntintDHLFields();
                                                DataTable ReferenciaSat = new DataTable();
                                                DataTable ReferenciaSat2 = new DataTable();


                                                ent.CP = e.ItemArray[tb_CP1].ToString();
                                                ent.CP = Convert.ToInt32(regexObj.Replace(ent.CP.Replace(" ", ""), "")).ToString(CPfmt);
                                                ent.CP2 = e.ItemArray[tb_CP2].ToString();
                                                ent.CP2 = Convert.ToInt32(regexObj.Replace(ent.CP2.Replace(" ", ""), "")).ToString(CPfmt);
                                                ReferenciaSat = ObtenerReferenciaSat(ent.CP);
                                                ReferenciaSat2 = ObtenerReferenciaSat(ent.CP2);

                                                ent.ReferenciaDelServicio = _fName;
                                                ent.RSdelRemitente = "SERVICIOS COMERCIALES AMAZON MEXICO S. DE R.L.";
                                                ent.RFCdelRemitente = "ANE140618P37";
                                                ent.Supplier = "ANE140618P37";
                                                ent.Calle = e.ItemArray[tb_Calle1].ToString();
                                                ent.Municipio = ReferenciaSat.Rows[0][2].ToString();
                                                ent.Estado = ReferenciaSat.Rows[0][1].ToString();
                                                ent.Pais = "MEX";
                                                ent.RSdelDestinatario = "SERVICIOS COMERCIALES AMAZON MEXICO S. DE R.L.";
                                                ent.RFCDestinatario = "ANE140618P37";
                                                ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                                ent.Municipio2 = ReferenciaSat2.Rows[0][2].ToString();
                                                ent.Estado2 = ReferenciaSat2.Rows[0][1].ToString();
                                                ent.Pais2 = "MEX";
                                                ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                                ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                                ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString().Trim(new Char[] { '\'' });
                                                ent.ClaveUnidadDeMedida = "KGM";
                                                ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString().Replace('|', ' ').Replace(',', ' ').Replace('°', ' ').Replace('-', ' ')
                                                    .Replace('\\', ' ').Replace('/', ' ').Replace('\'', ' ').Replace('’', ' ').Replace('"', ' ');
                                                ent.MaterialPeligroso = "No";
                                                ent.ValorDeLaMercancia = "1";
                                                ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                                if (ent.ReferenciaDelServicio != "")
                                                    eDhlList.Add(ent);

                                            }
                                        
                                            else
                                            {
                                                //Mandar email
                                                ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                                EnviarEmail(directories, fileInfo.Name, "Revisar VRID");
                                                break;
                                            }
                                        }
                                    }
                                    if (eDhlList.Any())
                                    {
                                        List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                        ClientList = VerificaMercanciasClientes(eDhlList);
                                        bool isEmpt = !ClientList.Any();
                                        if (isEmpt) { }
                                        else
                                        {
                                            InsertaMercanciaClientesList(ClientList, directories, _fNameInfo);
                                        }
                                    }
                                    //CortarDocumentos(directories);

                                }
                                else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                            }
                        }

                        //if (eDhlList.Any())
                        //{
                        //    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                        //    ClientList = VerificaMercanciasClientes(eDhlList);
                        //    bool isEmpt = !ClientList.Any();
                        //    if (isEmpt) { }
                        //    else
                        //    {
                        //        InsertaMercanciaClientesList(ClientList, directories, _fNameInfo);
                        //    }
                        //}

                        CortarDocumentos(directories);

                    }
                }
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);


                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Amazon...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void GetStellantis_Files()
        {

            string _fName = "";
            
            string strFullNme = "";


            int tb_EmbarqueDHL = 0;
            int tb_Orden = 0;
            int tb_IDOrigen = 0;
            //Sin valor
            int tb_RFCRemitente2 = 3;
            int tb_RSRemitente2 = 2;
            int tb_RSDest = 9;
            int tb_Calle = 4;
            int tb_Municipio = 5;
            int tb_Estado = 6;
            int tb_Pais = 7;
            int tb_CodigoPostal = 8;
            int tb_IDDestino = 0;
            //Sin valor
            int tb_RFCDestinatario2 = 10;
            //SinValor
            int tb_Calle2 = 11;
            int tb_Municipio3 = 12;
            int tb_Estado4 = 13;
            int tb_Pais5 = 14;
            int tb_CodigoPostal6 = 15;
            int tb_PesoNetoTotal = 16;
            int tb_NumTotalMercancias = 17;
            int tb_BienesTransp = 18;
            int tb_Descripción = 19;
            int tb_ClaveUnidad = 20;
            int tb_CveMaterialPeligroso = 21;
            int tb_Embalaje = 23;
            int tb_DescripEmbalaje = 24;
            int tb_ValorMercancia = 22;
            int tb_FraccionArancelaria = 23;
            int tb_UUIDComercioExt = 27;
            int tb_TotalKMRuta = 28;
            int InicioHeader = 1;
            int InicioTabla = 2;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();



            var directories = @"\\10.1.1.30\e$\Attachments\sftr-stellantis";
            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            bool direxists = System.IO.Directory.Exists(directories);


            
            try
            {
                // Setup session options
                

                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                if (direxists)
                {

                    FileInfo[] filesChek = (new DirectoryInfo(directories)).GetFiles();

                    if (filesChek.Length != 0)
                    {
                        FileInfo fileInfoCheck = filesChek[0];
                        _fName = fileInfoCheck.Name;
                        if ((fileInfoCheck.Name.Contains(".xlsx") ? true : fileInfoCheck.Name.Contains(".xlsx")))
                        {
                            _fName = _fName.ToString().Split('.')[0];
                        }


                        FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                        for (int j = 0; j < (int)files.Length; j++)
                        {
                            FileInfo fileInfo = files[j];

                            if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")) && (fileInfo.Name.Contains(_fName) ? true : fileInfo.Name.Contains(_fName)))
                            {
                                if ((fileInfoCheck.Name.Contains("-") ? true : fileInfoCheck.Name.Contains("-")))
                                {
                                    _fName = _fName.ToString().Split('-')[0];
                                }
                                String filenameToUpload = fileInfo.Name;
                                strFullNme = fileInfo.FullName;

                                if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                                {
                                    DataTable excelDT = new DataTable();
                                    excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                    errorTable = excelDT;
                                    if (excelDT.Rows.Count == 0)
                                    {
                                        ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                        EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                    }
                                    //El primer numero determina cuando comienzala hoja, despues  el header el tercer numero es donde comienza la tabla
                                    //1 = hoja, 6 = header, 8 = comienzo de la tabla


                                    foreach (DataRow e in excelDT.Rows)
                                    {
                                        try
                                        {
                                            if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_EmbarqueDHL].ToString() != "")
                                                && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                        && (excelDT.Columns["Calle"].Ordinal == tb_Calle))
                                            {

                                                EntintDHLFields ent = new EntintDHLFields();

                                                ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString();
                                                ent.RSdelRemitente = e.ItemArray[tb_RSRemitente2].ToString();
                                                ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                                ent.Supplier = e.ItemArray[tb_EmbarqueDHL].ToString();
                                                ent.Calle = e.ItemArray[tb_Calle].ToString();
                                                ent.Municipio = e.ItemArray[tb_Municipio].ToString();
                                                ent.Estado = e.ItemArray[tb_Estado].ToString();
                                                ent.Pais = e.ItemArray[tb_Pais].ToString();
                                                if (e.ItemArray[tb_CodigoPostal].ToString().Contains(' '))
                                                { ent.CP = e.ItemArray[tb_CodigoPostal].ToString().Split(' ')[1].Substring(0, 5); }
                                                else { ent.CP = e.ItemArray[tb_CodigoPostal].ToString().Substring(0, 5); }
                                                ent.RSdelDestinatario = e.ItemArray[tb_RSDest].ToString();
                                                ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                                ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                                ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString();
                                                ent.Estado2 = e.ItemArray[tb_Estado].ToString();
                                                ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                                if (e.ItemArray[tb_CodigoPostal6].ToString().Contains(' '))
                                                { ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString().Split(' ')[1].Substring(0, 5); }
                                                else { ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString().Substring(0, 5); }
                                                ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                                ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                                ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString().Trim(new Char[] { '\'' });
                                                ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString();
                                                ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString().Replace('\'', ' ');
                                                ent.MaterialPeligroso =e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                                ent.ValorDeLaMercancia =e.ItemArray[tb_ValorMercancia].ToString();
                                                ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                                if (ent.ReferenciaDelServicio != "")
                                                    eDhlList.Add(ent);
                                            }
                                            else
                                            {
                                                //Mandar email
                                                ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                                EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                                                break;
                                            }

                                        }
                                        catch (Exception)
                                        {
                                            ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                            EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                            throw;
                                        }

                                    }
                                }
                                else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                            }
                        }

                        if (eDhlList.Any())
                        {
                            List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                            ClientList = VerificaMercanciasClientes(eDhlList);
                            bool isEmpt = !ClientList.Any();
                            if (isEmpt) { }
                            else
                            {
                                InsertaMercanciaClientesList(ClientList, directories, _fName);
                            }
                        }
                        CortarDocumentos(directories);

                }
                    this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                    this.ltb_log.Items.Add("");
                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                }
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Stellantis...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void GetLear_Files()
        {

            string _fName = "";
            
            string strFullNme = "";
            string FaltaVaor = "FaltaValor";

            DateTime ClockInfoFromSystem = DateTime.Now;
            //int unidadEnv = 0;

            int tb_EmbarqueDHL = 0;
            int tb_Orden = 0;
            int tb_IDOrigen = 0;
            int tb_RFCRemitente2 = 0;
            //SinValor
            int tb_Calle = 0;
            int tb_Municipio = 0;
            //SinValor
            int tb_Estado = 0;
            int tb_Pais = 0;
            int tb_CodigoPostal = 0;
            int tb_IDDestino = 0;
            int tb_RFCDestinatario2 = 0;
            //SinValor
            int tb_Calle2 = 0;
            int tb_Municipio3 = 0;
            int tb_Estado4 = 0;
            int tb_Pais5 = 0;
            int tb_CodigoPostal6 = 0;
            int tb_PesoNetoTotal = 13;
            int tb_NumTotalMercancias = 10;
            int tb_Descripción = 0;
            int tb_BienesTransp = 7;
            int tb_ClaveUnidad = 0;
            int tb_CveMaterialPeligroso = 0;
            int tb_FraccionArancelaria = 15;
            int tb_ValorMercancia = 0;


            int tb_Embalaje = 0;
            int tb_DescripEmbalaje = 0;
            int tb_UUIDComercioExt = 0;
            int tb_TotalKMRuta = 0;
            int InicioHeader = 1;
            int InicioTabla = 2;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();

            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            var directories = @"\\10.1.1.30\e$\Attachments\sftr-learmexico";
            bool direxists = System.IO.Directory.Exists(directories);


            
            try
            {
                // Setup session options
                
                if (direxists)
                {

                    //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;

                            FileStream fileStream = null;
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

                            //var dateString = "12-12-2021 12:00:00";
                            //DateTime date1 = DateTime.Parse(dateString, System.Globalization.CultureInfo.CurrentCulture);

                            //DataTable excelDT = new DataTable();
                            //excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                            //errorTable = excelDT;
                            //if (excelDT.Rows.Count == 0)
                            //{
                            //    EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
                            //}
                            //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                            //6 = header, 8 = comienzo de la tabla
                            //if((int)ClockInfoFromSystem.DayOfWeek == 1)
                            //    unidadEnv = 28358;
                                
                            //if ((int)ClockInfoFromSystem.DayOfWeek == 3)
                            //    unidadEnv = 28224;

                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && (!_fName.Contains("~$")))
                            {
                                DataTable cargaRS = new DataTable();
                                DataTable origenCarg = new DataTable();
                                DataTable destinoCarg = new DataTable();
                                DataTable ReferenciaSat = new DataTable();

                                //cargaRS = ObtenerCargaPenske(fileInfo.Name.ToString().Split('_')[1].ToString().Split('.')[0]);
                                cargaRS = ObtenerCargaPenske(fileInfo.Name.ToString().Split('.')[0]);
                                origenCarg = ObtenerOrigenCarga(Int32.Parse(cargaRS.Rows[0][3].ToString()));
                                destinoCarg = ObtenerDestinoCarga(Int32.Parse(cargaRS.Rows[0][3].ToString()));
                                //ReferenciaSat = ObtenerReferenciaSat(origenCarg.Rows[0][3].ToString());


                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                int rows = worksheet.Dimension.Rows;
                                for (int e = 2; e < rows; e++)
                                {
                                    try
                                    {
                                        if ((worksheet.Cells[e, tb_NumTotalMercancias].Value?.ToString() != null) && (worksheet.Cells[e, tb_NumTotalMercancias].Value?.ToString() != "")
                                            && ((worksheet.Cells[e, tb_BienesTransp].Value?.ToString() != null) && (worksheet.Cells[e, tb_BienesTransp].Value?.ToString() != "")))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            ent.ReferenciaDelServicio = cargaRS.Rows[0][2].ToString();
                                            ent.RSdelRemitente = origenCarg.Rows[0][0].ToString();
                                            ent.RFCdelRemitente = origenCarg.Rows[0][7].ToString();
                                            ent.Supplier = "0000";
                                            ent.Calle = origenCarg.Rows[0][1].ToString();
                                            ent.Municipio = origenCarg.Rows[0][11].ToString();
                                            ent.Estado = origenCarg.Rows[0][5].ToString();
                                            ent.Pais = "MEX";
                                            ent.CP = origenCarg.Rows[0][3].ToString();
                                            ent.RSdelDestinatario = destinoCarg.Rows[0][0].ToString();
                                            ent.RFCDestinatario = destinoCarg.Rows[0][7].ToString();
                                            ent.Calle2 = destinoCarg.Rows[0][1].ToString();
                                            ent.Municipio2 = destinoCarg.Rows[0][11].ToString();
                                            ent.Estado2 = destinoCarg.Rows[0][5].ToString();
                                            ent.Pais2 = "MEX";
                                            ent.CP2 = destinoCarg.Rows[0][3].ToString();
                                            ent.PesoNeto = worksheet.Cells[e, tb_PesoNetoTotal].Value?.ToString();
                                            ent.NumeroTotalMercancias = worksheet.Cells[e, tb_NumTotalMercancias].Value?.ToString();
                                            ent.ClaveDelBienTransportado = worksheet.Cells[e, tb_BienesTransp].Value?.ToString();
                                            ent.ClaveUnidadDeMedida = "KGM";
                                            ent.DescripcionDelBienTransportado = worksheet.Cells[e, tb_BienesTransp].Value?.ToString();
                                            ent.MaterialPeligroso = "No";
                                            ent.ValorDeLaMercancia = "1";
                                            ent.TipoDeMoneda = "MXN";

                                            if (ent.PesoNeto != null)
                                                eDhlList.Add(ent);
                                            else
                                                break;

                                        }
                                        else
                                        {
                                            break;
                                        }

                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }

                                }
                                if (eDhlList.Any())
                                {
                                    List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                    ClientList = VerificaMercanciasClientes(eDhlList);
                                    bool isEmpt = !ClientList.Any();
                                    if (isEmpt) { }
                                    else
                                    {
                                        InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }
                        }
                    }

                    CortarDocumentos(directories);

                }
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));
                
                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Lear...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void GetTremec_Files()
        {

            string _fName = "";

            string strFullNme = "";

            int tb_EmbarqueDHL = 0;
            int tb_Orden = 1;
            int tb_RSremitente = 1;
            int tb_IDOrigen = 1;
            int tb_RFCRemitente2 = 2;
            int tb_Calle = 3;
            int tb_Municipio = 4;
            int tb_Estado = 5;
            int tb_Pais = 6;
            int tb_CodigoPostal = 7;
            int tb_IDDestino = 10;
            int tb_RSDest = 8;
            int tb_RFCDestinatario2 = 9;
            int tb_Calle2 = 10;
            int tb_Municipio3 = 11;
            int tb_Estado4 = 12;
            int tb_Pais5 = 13;
            int tb_CodigoPostal6 = 14;
            int tb_PesoNetoTotal = 15;
            int tb_NumTotalMercancias = 16;
            int tb_BienesTransp = 17;
            int tb_Descripción = 18;
            int tb_ClaveUnidad = 19;
            int tb_CveMaterialPeligroso = 20;
            int tb_Embalaje = 22;
            int tb_DescripEmbalaje = 23;
            int tb_ValorMercancia = 21;
            int tb_FraccionArancelaria = 22;
            int tb_UUIDComercioExt = 26;
            int tb_TotalKMRuta = 27;
            int InicioHeader = 1;
            int InicioTabla = 3;


            int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
            DataTable errorTable = new DataTable();



            //var directories = @"C:\Users\eanavia\Desktop\FTRDHL\Formatos\Test";
            var directories = @"\\10.1.1.30\e$\Attachments\sftr-tremec";
            bool direxists = System.IO.Directory.Exists(directories);

            try
            {
                // Setup session options


                if (direxists)
                {

                    FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
                    for (int j = 0; j < (int)files.Length; j++)
                    {
                        FileInfo fileInfo = files[j];
                        if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
                        {

                            String filenameToUpload = fileInfo.Name;
                            _fName = fileInfo.Name;
                            strFullNme = fileInfo.FullName;
                            
                            if ("NO" == EstaOnoBaneado(fileInfo.FullName) && ( ! _fName.Contains("~$")))
                            {

                                DataTable excelDT = new DataTable();
                                excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, InicioHeader, InicioTabla);
                                errorTable = excelDT;

                                //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
                                //6 = header, 8 = comienzo de la tabla

                                if (excelDT.Rows.Count == 0)
                                {
                                    ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
                                    EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
                                }

                                List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

                                foreach (DataRow e in excelDT.Rows)
                                {
                                    try
                                    {
                                        if (((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && ((e.ItemArray[tb_EmbarqueDHL].ToString() != "")
                                            && (e.ItemArray[tb_RFCDestinatario2].ToString() != null)) && (e.ItemArray[tb_RFCDestinatario2].ToString() != ""))
                                            && ((e.ItemArray[tb_BienesTransp].ToString() != null) && (e.ItemArray[tb_BienesTransp].ToString() != ""))
                                                && (excelDT.Columns["Calle"].Ordinal == tb_Calle) && (excelDT.Columns["Estado"].Ordinal == tb_Estado))
                                        {

                                            EntintDHLFields ent = new EntintDHLFields();

                                            if (e.ItemArray[tb_EmbarqueDHL].ToString().Contains('.'))
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString().Split('.')[1]; }
                                            else
                                            { ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString(); }
                                            ent.RSdelRemitente = e.ItemArray[tb_RSremitente].ToString();
                                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
                                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
                                            ent.Calle = e.ItemArray[tb_Calle].ToString();
                                            ent.Municipio = e.ItemArray[tb_Municipio].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado = e.ItemArray[tb_Estado].ToString();
                                            ent.Pais = e.ItemArray[tb_Pais].ToString();
                                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString().Trim(new Char[] { '\'' });
                                            ent.RSdelDestinatario = e.ItemArray[tb_RSDest].ToString();
                                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
                                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
                                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString().Trim(new Char[] { '\'' });
                                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString().Trim(new Char[] { '\'' });
                                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
                                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString().Trim(new Char[] { '\'' });
                                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
                                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
                                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
                                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString().Trim(new Char[] { '\'' });
                                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
                                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
                                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
                                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

                                            if (ent.ReferenciaDelServicio != "")
                                                eDhlList.Add(ent);

                                        }
                                        else
                                        {
                                            //Mandar email
                                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla O se ocultaron columnas{0}", "<br />");
                                            EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                            break;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        ErrorMessageTxt += String.Format("El orden de filas no es el correcto o no esta bien escrito O se ocultaron Columnas {0}", "<br />");
                                        EnviarEmail(directories, _fName, DetectarErrorFormato(errorTable, ordenColm));
                                        throw;
                                    }
                                }
                                if (eDhlList.Any())
                                {
                                    {
                                        List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
                                        ClientList = VerificaMercanciasClientes(eDhlList);
                                        bool isEmpt = !ClientList.Any();
                                        if (isEmpt) { }
                                        else
                                        {
                                            InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
                                        }
                                    }
                                }
                            }
                            else { this.ltb_log.Items.Add(string.Format(" El archivo se encuentra en uso... {0}", DateTime.Now)); }


                        }
                    }

                    CortarDocumentos(directories);

                }
                //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
                this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
            }
            catch (Exception ex)
            {
                string erro = ex.ToString();
                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
                ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                EnviarEmail(directories, _fName, ErrorMessageTxt);

                CortarDocumentosPError(directories, _fName);
                //ActualizaTransaccionEnviado(_fName.Substring(7, 11));

                this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion GenericF...", DateTime.Now));
                this.ltb_log.Items.Add("");
                //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
                //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void CortarDocumentos(string pathDir)
        {
            DateTime moment = DateTime.Now;
            string monthDate = moment.Month +"-"+moment.Year;
     
            bool direxists = System.IO.Directory.Exists(@"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(pathDir).Name +"\\"+ monthDate);

            if (!direxists)
                System.IO.Directory.CreateDirectory(@"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(pathDir).Name + "\\" + monthDate);

            string fileCutName = "";
            string sourcePath = pathDir;
            string targetPath = @"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(pathDir).Name + "\\" + monthDate;

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

                            Thread.Sleep(1000);
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
        private string LimpiaXMLDocs(string pathDir)
        {
            StreamReader streamReader = new StreamReader(pathDir);
            string text = streamReader.ReadToEnd();
            streamReader.Close();
            return text;
        }

        public string RemoveSpecialCharacters(string str)
        {
            //change regular expression as per your need
            byte[] tempBytes = System.Text.Encoding.GetEncoding("ISO-8859-8").GetBytes(str);
            string cleanText = System.Text.Encoding.UTF8.GetString(tempBytes);
            return cleanText;
        }

        public void SaveXmlCleanDoc(string xmlstring, string pathDir)
        {
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xmlstring);
            doc.PreserveWhitespace = true;
            File.Delete(pathDir);
            doc.Save(pathDir);
        }

        private void CortarDocumentosXML(string pathDir)
        {
            DateTime moment = DateTime.Now;
            string monthDate = moment.Month + "-" + moment.Year;

            bool direxists = System.IO.Directory.Exists(@"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(pathDir).Name + "\\" + monthDate);

            if (!direxists)
                System.IO.Directory.CreateDirectory(@"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(pathDir).Name + "\\" + monthDate);

            string fileCutName = "";
            string sourcePath = pathDir;
            string targetPath = @"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(pathDir).Name + "\\" + monthDate;

            // Use Path class to manipulate file and directory paths.
            string destFile = System.IO.Path.Combine(targetPath, fileCutName);

            if (System.IO.Directory.Exists(sourcePath))
            {
                string[] filesIter = System.IO.Directory.GetFiles(sourcePath);

                // Copy the files Iterand overwrite destination files Iterif they already exist.
                foreach (string s in filesIter)
                {
                    fileCutName = System.IO.Path.GetFileName(s);
                    
                    // Use static Path methods to extract only the file name from the path.
                    destFile = System.IO.Path.Combine(targetPath, fileCutName);
                    System.IO.File.Copy(s, destFile, true);
                    try
                    {

                        Thread.Sleep(1000);
                        System.IO.File.Delete(s);
                    }
                    catch (System.IO.IOException e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
            }
        }

        private void CortarDocumentosPError(string pathDir, string filName)
        {
            DateTime moment = DateTime.Now;
            string monthDate = moment.Month + "-" + moment.Year;

            bool direxists = System.IO.Directory.Exists(@"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(pathDir).Name + "\\" + monthDate);

            if (!direxists)
                System.IO.Directory.CreateDirectory(@"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(pathDir).Name + "\\" + monthDate);

            string fileCutName = "";
            string sourcePath = pathDir;
            string targetPath = @"\\10.1.1.30\FTProot\Done\" + new DirectoryInfo(pathDir).Name + "\\" + monthDate;

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
                        if ((fileCutName.Contains(filName) ? true : fileCutName.Contains(filName)))
                        {
                            // Use static Path methods to extract only the file name from the path.
                            destFile = System.IO.Path.Combine(targetPath, fileCutName);
                            System.IO.File.Copy(s, destFile, true);
                            try
                            {
                                //Process proc = Process.GetProcessesByName(s);
                                //// do stuff
                                //proc.Kill();
                                
                                Thread.Sleep(1000);
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
        }




        private DataTable ProcesoExcel(DataTable excelDTF, string excelPath,int dataShet, int strow, int startfor)
        {
            bool isBanned = false;
            for (int i = 0; i < BannedExcel.Count; i++)
            {
                if (excelPath == BannedExcel[i]) 
                {isBanned = true;}
            }
            if (isBanned)
            {
            }
            else
            {
                try
                {
                    FileStream fileStream = null;

                    // limpieza ante todo
                    excelDTF.Clear();

                    FileInfo datos = new FileInfo(excelPath);
                    ExcelPackage LecExcel = new ExcelPackage(datos);
                    try
                    {
                        fileStream = datos.Open(FileMode.Open, FileAccess.ReadWrite);
                    }
                    catch (Exception)
                    {
                        BannedExcel.Add(excelPath);
                        //~$
                        ErrorMessageTxt += String.Format("El archivo no pudo leerse, Cierre el archivo para que pueda procesarse {0}", "<br />");
                        return excelDTF;
                    }
                    //catch (NullReferenceException e)
                    //{
                    //    MessageBox.Show("El archivo se encuentra abierto por otro proceso {0}", e.ToString());
                    //    return excelDTF;
                    //}
                    //catch (IOException ex)
                    //{
                    //    BannedExcel.Add(excelPath);
                    //    MessageBox.Show("El archivo se encuentra abierto por otro proceso. {0}", ex.ToString());
                    //    return excelDTF;
                    //}
                    LecExcel.Load(fileStream);

                    try
                    {
                        ExcelWorksheet worksheetTest = LecExcel.Workbook.Worksheets[dataShet];
                        fileStream.Dispose();
                        fileStream.Close();

                    }
                    catch (Exception)
                    {
                        fileStream.Dispose();
                        fileStream.Close();
                        ErrorMessageTxt += String.Format("El archivo no pudo leerse, puede que la hoja de Excel no sea la correcta{0}", "<br />");
                        throw;
                    }

                    ExcelWorksheet worksheet = LecExcel.Workbook.Worksheets[dataShet];
                    fileStream.Dispose();
                    fileStream.Close();


                    if (worksheet.Dimension == null)
                    {
                        ErrorMessageTxt += String.Format("La hoja esta vacía {0}", "<br />");
                        return excelDTF;
                    }

                    //create a list to hold the column names
                    List<string> columnNames = new List<string>();

                    //needed to keep track of empty column headers
                    int currentColumn = 1;

                    //loop all columns in the sheet and add them to the datatable
                    //Lugar 1 y 2 van a ser cambiados por variables el primero es row y el seg, es colm
                    foreach (var cell in worksheet.Cells[strow, 1, 1, worksheet.Dimension.End.Column])
                    {
                        string columnName = cell.Text.Trim();

                        //check if the previous header was empty and add it if it was
                        //if (cell.Start.Column != currentColumn)
                        //{
                        //    columnNames.Add("Header_" + currentColumn);
                        //    excelDTF.Columns.Add("Header_" + currentColumn);
                        //    currentColumn++;
                        //}

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
                        excelDTF.Columns.Add(columnName);

                        currentColumn++;
                    }

                    for (int i = startfor; i <= worksheet.Dimension.End.Row; i++)
                    {
                        var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                        DataRow newRow = excelDTF.NewRow();
             
                        //loop all cells in the row
                        foreach (var cell in row)
                        {
                            if (cell.Text.Contains("*NOTA:") || cell.Text.Contains("*Nota:") || cell.Text.Contains("nota:"))
                            {

                            }
                            else { newRow[cell.Start.Column - 1] = CleanInput(cell.Text).Trim(); }
                        }
                        if((newRow.ItemArray[0].ToString() == "") && (newRow.ItemArray[1].ToString() == "") && (newRow.ItemArray[2].ToString() == "") && (newRow.ItemArray[4].ToString() == "")) 
                            i = i + worksheet.Dimension.End.Row;
                        else { excelDTF.Rows.Add(newRow); }
                    }
                    fileStream.Dispose();

                    return excelDTF;

                }
                catch (Exception)
                {
                    ErrorMessageTxt += String.Format("El archivo no pudo leerse, puede que la hoja de Excel no sea la correcta{0}", "<br />");
                    throw;
                }
            }
            isBanned = false;
            return excelDTF;
        }


        public string DetectarErrorFormato(DataTable excelDTF, int[] arr)
        {
            String mensaje = "";
            int tb_EmbarqueDHL = arr[0];
            int tb_Orden = arr[1];
            int tb_IDOrigen = arr[2];
            int tb_RFCRemitente2 = arr[3];
            int tb_Calle = arr[4];
            int tb_Municipio = arr[5];
            int tb_Estado = arr[6];
            int tb_Pais = arr[7];
            int tb_CodigoPostal = arr[8];
            int tb_IDDestino = arr[9];
            int tb_RFCDestinatario2 = arr[10];
            int tb_Calle2 = arr[11];
            int tb_Municipio3 = arr[12];
            int tb_Estado4 = arr[13];
            int tb_Pais5 = arr[14];
            int tb_CodigoPostal6 = arr[15];
            int tb_PesoNetoTotal = arr[16];
            int tb_NumTotalMercancias = arr[17];
            int tb_BienesTransp = arr[18];
            int tb_Descripción = arr[19];
            int tb_ClaveUnidad = arr[20];
            int tb_CveMaterialPeligroso = arr[21];
            int tb_Embalaje = arr[22];
            int tb_DescripEmbalaje = arr[23];
            int tb_ValorMercancia = arr[24];
            int tb_FraccionArancelaria = arr[25];
            int tb_UUIDComercioExt = arr[26];
            int tb_TotalKMRuta = arr[27];
            int InicioHeader = arr[28];
            int InicioTabla = arr[29];

            int maxValue = arr.Max();

            try
            {
                mensaje += ErrorMessageTxt;
                //if ((excelDTF.Columns.Count) != maxValue)
                //{
                //    mensaje += String.Format("El numero de Columnas no coincide con el que se espera, Se enviaron {1} Se esperan {2}:{0}", "<br />", excelDTF.Columns.Count.ToString(), maxValue.ToString());
                //}
                mensaje += String.Format("Es posible que la tabla empiece en una posicion diferente a la acordada{0} Inicio de Headers En Fila {1} Inicio de tabla en {2} {0}", "<br />", InicioHeader, InicioTabla);
                //mensaje += String.Format("Estos pueden ser los sig motivos por el cual el formato no se esta procesando:{0}", "<br />");

                if (excelDTF != null)
                {
                    for (int i = 0; i < excelDTF.Columns.Count-1; i++)
                    {
                        if ((excelDTF.Columns[i].ColumnName.Contains("REQ"))|| (excelDTF.Columns[i].ColumnName.Contains("Req"))|| (excelDTF.Columns[i].ColumnName.Contains("req")) || (excelDTF.Columns[i].ColumnName.Contains("CLI")))
                            mensaje += String.Format("La columna de headers se movio :{0}", "<br />");
                        if ((excelDTF.Rows[0].ItemArray[i].ToString().Contains("REQ")) || (excelDTF.Rows[0].ItemArray[i].ToString().Contains("Req")) || (excelDTF.Rows[0].ItemArray[i].ToString().Contains("req")) || (excelDTF.Rows[0].ItemArray[i].ToString().Contains("CLI")))
                            mensaje += String.Format("La columna de No concuerda favor de quitar la colm de requerido se movio :{0}", "<br />");

                    }

                    if (!((excelDTF.Columns[tb_EmbarqueDHL].ColumnName.Contains("Ref")) ||(excelDTF.Columns[tb_EmbarqueDHL].ColumnName.Contains("EMB")) || (excelDTF.Columns[tb_EmbarqueDHL].ColumnName.Contains("JB"))))
                    { mensaje += String.Format("El nombre de la columna que contiene la Referencia del servicio esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_EmbarqueDHL); 
                        return mensaje;}
                    if (!excelDTF.Columns[tb_RFCRemitente2].ColumnName.Contains("RFC"))
                        {mensaje += String.Format("El nombre de la columna que contiene la RFC esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_RFCRemitente2);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_Calle].ColumnName.Contains("Calle"))||(excelDTF.Columns[tb_Calle].ColumnName.Contains("CALLE")) ||(excelDTF.Columns[tb_Calle].ColumnName.Contains("Dom"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Calle esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_Calle);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_Municipio].ColumnName.Contains("Mun"))||(excelDTF.Columns[tb_Municipio].ColumnName.Contains("MUN")) ||(excelDTF.Columns[tb_Municipio].ColumnName.Contains("ID"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Municipio esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_Municipio);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_Estado].ColumnName.Contains("Est"))||(excelDTF.Columns[tb_Estado].ColumnName.Contains("EST"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Estado esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_Estado);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_Pais].ColumnName.Contains("Pa"))||(excelDTF.Columns[tb_Pais].ColumnName.Contains("PA"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Pais esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_Pais);
                        return mensaje;}
                    if (!excelDTF.Columns[tb_CodigoPostal].ColumnName.Contains("C"))
                        {mensaje += String.Format("El nombre de la columna que contiene la CodigoPostal esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_CodigoPostal);
                        return mensaje;}
                    if (!excelDTF.Columns[tb_RFCDestinatario2].ColumnName.Contains("RFC"))
                        {mensaje += String.Format("El nombre de la columna que contiene la RFC esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_RFCDestinatario2);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_Calle2].ColumnName.Contains("Calle"))||(excelDTF.Columns[tb_Calle2].ColumnName.Contains("CALLE")) ||(excelDTF.Columns[tb_Calle2].ColumnName.Contains("Dom"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Calle esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_Calle2);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_Municipio3].ColumnName.Contains("Mun"))||(excelDTF.Columns[tb_Municipio3].ColumnName.Contains("MUN")) ||(excelDTF.Columns[tb_Municipio3].ColumnName.Contains("ID"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Mun esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_Municipio3);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_Estado4].ColumnName.Contains("Est"))||(excelDTF.Columns[tb_Estado4].ColumnName.Contains("EST")) ||(excelDTF.Columns[tb_Estado4].ColumnName.Contains("ID"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Estado esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_Estado4);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_Pais5].ColumnName.Contains("Pa"))||(excelDTF.Columns[tb_Pais5].ColumnName.Contains("PA")) ||(excelDTF.Columns[tb_Pais5].ColumnName.Contains("ID"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Pais esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_Pais5);
                        return mensaje;}
                    if (!excelDTF.Columns[tb_CodigoPostal6].ColumnName.Contains("C"))
                        {mensaje += String.Format("El nombre de la columna que contiene la CP esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_CodigoPostal6);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_PesoNetoTotal].ColumnName.Contains("Peso"))||(excelDTF.Columns[tb_PesoNetoTotal].ColumnName.Contains("PESO"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Peso esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_PesoNetoTotal);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_NumTotalMercancias].ColumnName.Contains("Total"))||(excelDTF.Columns[tb_NumTotalMercancias].ColumnName.Contains("TOT")) ||(excelDTF.Columns[tb_NumTotalMercancias].ColumnName.Contains("Num"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Numero Total de Mercancias esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_NumTotalMercancias);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_BienesTransp].ColumnName.Contains("Bien"))||(excelDTF.Columns[tb_BienesTransp].ColumnName.Contains("BIEN")) ||(excelDTF.Columns[tb_BienesTransp].ColumnName.Contains("Cv"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Clave Bienes Transp esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_BienesTransp);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_Descripción].ColumnName.Contains("Desc"))||(excelDTF.Columns[tb_Descripción].ColumnName.Contains("DESC"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Descripcion esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_Descripción);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_ClaveUnidad].ColumnName.Contains("Clave"))||(excelDTF.Columns[tb_ClaveUnidad].ColumnName.Contains("CLA")) ||(excelDTF.Columns[tb_ClaveUnidad].ColumnName.Contains("Uni"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Clave Unidad esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_ClaveUnidad);
                        return mensaje;}
                    if (!((excelDTF.Columns[tb_CveMaterialPeligroso].ColumnName.Contains("Material"))||(excelDTF.Columns[tb_CveMaterialPeligroso].ColumnName.Contains("MAT"))))
                        {mensaje += String.Format("El nombre de la columna que contiene la Material Peligroso esta mal escrita o en una columna diferente, deberia estar en {1} {0}", "<br />", tb_CveMaterialPeligroso);
                        return mensaje;}

                    //TODO : Limpiar el mensaje de error para el correo cada vez que se llame a una funcion para leer un excel, al principio de todo  
                    //TODO : hacer correcciones en llos if de abajo 
                    //for (int e = 2; e < excelDTF.Rows; e++)
                    //if ((excelDTF.Rows[0].ItemArray[tb_EmbarqueDHL].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_EmbarqueDHL].ToString() != null))
                    //Es solo cambiar la condicion de != a == 
                    //Y meter el if en un for 
                    //{ }

                    if ((excelDTF.Rows[0].ItemArray[tb_EmbarqueDHL].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_EmbarqueDHL].ToString() != null))
                        {mensaje += String.Format("El contenido de Referencia del servicio es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_RFCRemitente2].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_RFCRemitente2].ToString() != null))
                        {mensaje += String.Format("El contenido de RFC es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_Calle].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_Calle].ToString() != null))
                        {mensaje += String.Format("El contenido de Calle es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_Municipio].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_Municipio].ToString() != null))
                        {mensaje += String.Format("El contenido de Municipio es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_Estado].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_Estado].ToString() != null))
                        {mensaje += String.Format("El contenido de Estado es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_Pais].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_Pais].ToString() != null))
                        {mensaje += String.Format("El contenido de Pais es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_CodigoPostal].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_CodigoPostal].ToString() != null))
                        {mensaje += String.Format("El contenido de CodigoPostal es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_RFCDestinatario2].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_RFCDestinatario2].ToString() != null))
                        {mensaje += String.Format("El contenido de RFC es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_Calle2].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_Calle2].ToString() != null))
                        {mensaje += String.Format("El contenido de Calle es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_Municipio3].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_Municipio3].ToString() != null))
                        {mensaje += String.Format("El contenido de Mun es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_Estado4].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_Estado4].ToString() != null))
                        {mensaje += String.Format("El contenido de Estado es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_Pais5].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_Pais5].ToString() != null))
                        {mensaje += String.Format("El contenido de Pais es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_CodigoPostal6].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_CodigoPostal6].ToString() != null))
                        {mensaje += String.Format("El contenido de CP es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_PesoNetoTotal].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_PesoNetoTotal].ToString() != null))
                        {mensaje += String.Format("El contenido de Peso es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_NumTotalMercancias].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_NumTotalMercancias].ToString() != null))
                        {mensaje += String.Format("El contenido de Numero Total de Mercancias es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_BienesTransp].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_BienesTransp].ToString() != null))
                        {mensaje += String.Format("El contenido de Clave Bienes Transp es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_Descripción].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_Descripción].ToString() != null))
                        {mensaje += String.Format("El contenido de Descripcion es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_ClaveUnidad].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_ClaveUnidad].ToString() != null))
                        {mensaje += String.Format("El contenido de Clave Unidad es nulo {0}", "<br />");
                        return mensaje;}
                    if ((excelDTF.Rows[0].ItemArray[tb_CveMaterialPeligroso].ToString()!= "")|| (excelDTF.Rows[0].ItemArray[tb_CveMaterialPeligroso].ToString() != null))
                        {mensaje += String.Format("El contenido de Material Peligroso es nulo {0}", "<br />");
                        return mensaje;}
                }
                else
                {
                    mensaje += String.Format("Algo salio muy mal y no se pudo leer la tabla {0}","<br />");
                }

            }
            catch (Exception ex)
            {
                string errortxt = ex.ToString();
                if (ex.ToString().Contains("en "))
                {
                    errortxt = ex.ToString().Split('_')[0];
                }
                mensaje += String.Format("El El proceso se detuvo porque la tabla proporcionada es mas grande o mas pequeña  {0}", "<br />");

                mensaje += String.Format("El error es  {0} {1} Las Colmnas mximas son {2} Las columnas que de el formato rechazado son {4} {1} {3}","El formato Que se proceso tiene mas columnas que las acordadas", "<br />", maxValue.ToString() , errortxt, excelDTF.Columns.Count.ToString());
                return mensaje;

            }

            return mensaje;

        }


        public void EnviarEmail(string dest, string nombreArch, string mensaje)
        {

            string destEmail = "";
            string msgMail1 = "El archivo (" + nombreArch +") que se proporciono no cumple con los campos requeridos.";
            string msgMail2 = mensaje;
            string msgMail3 = "Contacto javega@ftr.com.mx ";
            string asunto = "Archivo no cumple con el orden o campos requeridos carpeta:";
            var origPath = dest +"\\"+ nombreArch;


            if (dest.Contains("adient"))
            {
                asunto += "adient";
            }
            if (dest.Contains("generico"))
            {
                asunto += "generico";
            }
            if (dest.Contains("amazon"))
            {
                asunto += "amazon";
            }
            if (dest.Contains("android"))
            {
                asunto += "android";
            }
            if (dest.Contains("apl"))
            {
                asunto += "apl";
            }
            if (dest.Contains("celtic"))
            {
                asunto += "celtic";
            }
            if (dest.Contains("dhlsupply"))
            {
                asunto += "dhlsupply";
            }
            if (dest.Contains("jbhunt"))
            {
                asunto += "jbhunt";
            }
            if (dest.Contains("leartoluca"))
            {
                asunto += "leartoluca";
                ErrorMessageTxt += String.Format("La ruta tiene que estar al final del nombre de el archivo precedido por un _ {0}", "<br />");
            }
            if (dest.Contains("matson"))
            {
                asunto += "matson";
            }
            if (dest.Contains("schneider"))
            {
                asunto += "schneider";
            }
            if (dest.Contains("stellantis"))
            {
                asunto += "stellantis";
            }
            if (dest.Contains("stellantis-mts"))
            {
                asunto += "stellantis-mts";
            }
            if (dest.Contains("transplace"))
            {
                asunto += "transplace";
            }
            if (dest.Contains("tremec"))
            {
                asunto += "tremec";
            }
            if (dest.Contains("truper"))
            {
                asunto += "truper";
            }
            if (dest.Contains("upds"))
            {
                asunto += "upds";
            }
            if (dest.Contains("Test"))
            {
                asunto += "test";
            }
            if (dest.Contains("learmexico"))
            {
                asunto += "learmexico";
                ErrorMessageTxt += String.Format("La ruta tiene que estar al final del nombre de el archivo precedido por un _ {0}", "<br />");
            }


            //destEmail = "desaftr02@ftr.com.mx";
            destEmail = "eanavia@ftr.com.mx";
            //Cambiar a true            //Cambiar a true            //Cambiar a true
            SendMail smail = new SendMail(true);
            smail.Notificar(destEmail, msgMail1, msgMail2, msgMail3, asunto, origPath);
            ErrorMessageTxt = "";
        }

        /// <summary>
        /// Procesa el archivo de excel y lo carga en el Datatable
        /// </summary>
        /// <param name="dt"></param>


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
                    //this.UploadFiles();
                    this.GetClientFiles_Message();
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
        //private void DirectorioDescarga()
        //{
        //    string str = DateTime.Now.ToString("MMMM");
        //    DateTime now = DateTime.Now;
        //    string str1 = string.Format("{0}{1}", str, now.Year);
        //    char upper = char.ToUpper(str1[0]);
        //    str1 = string.Concat(upper.ToString(), str1.Substring(1));
        //    this._xFolderPath = string.Format("C:\\ALTOMOVUP\\{0}", str1);
        //    this.di = new DirectoryInfo(this._xFolderPath);

        //    if (!Directory.Exists("C:\\ALTOMOVUP" + "\\ArchivosAF"))
        //    {
        //        // SE CREAN SUB CARPETAS PARA IDENTIFICAR LOS ARCHIVOS X6
        //        this.di = Directory.CreateDirectory("C:\\ALTOMOVUP" + "\\ArchivosAF");
        //        this.ltb_log.Items.Add(string.Format("{0} - Carpeta creada: {1}", DateTime.Now, "C:\\ALTOMOVUP" + "\\ArchivosAF"));
        //        this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //    }

        //    if (!Directory.Exists("C:\\ALTOMOVUP" + "\\ArchivosAFEnviado"))
        //    {
        //        this.di = Directory.CreateDirectory("C:\\ALTOMOVUP" + "\\ArchivosAFEnviado");
        //        this.ltb_log.Items.Add(string.Format("{0} - Carpeta creada: {1}", DateTime.Now, "C:\\ALTOMOVUP" + "\\ArchivosAFEnviado"));
        //        this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //    }
        //}




        //    private void GetFiles()
        //    {
        //        string _fName = "";
        //        try
        //        {
        //            DataTable dt = new DataTable();

        //            SqlCommand sqlCommand = new SqlCommand("SP_Archivos214_20_AltoMovup", this._xConnString)
        //            {
        //                CommandType = CommandType.StoredProcedure
        //            };

        //            this._xConnString.Open();

        //            SqlDataAdapter da = new SqlDataAdapter(sqlCommand);
        //            da.Fill(dt);

        //            this.ltb_log.Items.Add(string.Format("{0} - Transacciones Cargadas Correctamente..", DateTime.Now));
        //            this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //            for (int i = 0; i < dt.Rows.Count; i++)
        //            {
        //                _fName = dt.Rows[i]["ARCHIVO"].ToString() + ".txt";
        //                //Dado el caso, verifico que exista el archivo..
        //                if (!File.Exists(Constantes.SourceDir + _fName))
        //                {
        //                    File.Copy(Path.Combine(Constantes.SourceDir, _fName), Path.Combine("C:\\ALTOMOVUP" + "\\ArchivosAF", _fName), true);
        //                    this.ltb_log.Items.Add(string.Format("{0} - Archivo Procesado: {1}.", DateTime.Now, _fName));
        //                    this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //                }
        //            }
        //            this._xConnString.Close();
        //        }
        //        catch
        //        {
        //            this._xConnString.Close();
        //            //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
        ////            ActualizaTransaccionEnviado(_fName.Substring(7, 11));
        //            this.ltb_log.Items.Add(string.Format("{0} - Error al cargar los Archivos.", DateTime.Now));
        //            this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //        }
        //    }
        //    private void UploadFiles()
        //    {

        //        string _fName = "";
        //        
        //        string strFullNme = "";
        //        
        //        try
        //        {
        //            // Setup session options
        //            SessionOptions sessionOptions = new SessionOptions
        //            {
        //                Protocol = Protocol.Ftp,
        //                FtpSecure= FtpSecure.Explicit,
        //                FtpMode = FtpMode.Passive,
        //                PortNumber = Constantes.PuertoSFTP,
        //                HostName = Constantes.ServidorSFTP, //this.txtHostName.Text,
        //                UserName = Constantes.UsuarioSFTP, //this.txtUserName.Text,
        //                Password = Constantes.PassWordSFTP, //this.txtPwd.Text,
        //                GiveUpSecurityAndAcceptAnySshHostKey = false,
        //                GiveUpSecurityAndAcceptAnyTlsHostCertificate = true
        //            };

        //            session.Open(sessionOptions);
        //            FileInfo[] files = (new DirectoryInfo(@"C:\ALTOMOVUP\\ArchivosAF")).GetFiles();
        //            for (int i = 0; i < (int)files.Length; i++)
        //            {
        //                FileInfo fileInfo = files[i];
        //                str = string.Concat(@"C:\ALTOMOVUP\ArchivosAFEnviado\", fileInfo.Name);
        //                if ((fileInfo.Name.Contains(".txt") ? true : fileInfo.Name.Contains(".txt")))
        //                {
        //                    String RemotePath = Constantes.RutaUploadSFTP; //this.txtRutaSubidaSFTP.Text; 
        //                    String filenameToUpload = fileInfo.Name;
        //                    String localPathtoUpload = @"C:\ALTOMOVUP\ArchivosAF\"; //this.txtRutaSubidaLocal.Text;
        //                    _fName = fileInfo.Name;
        //                    strFullNme = fileInfo.FullName;
        //                    if (session.Opened)
        //                    {
        //                        //Get Ftp File
        //                        TransferOptions transferOptions = new TransferOptions();
        //                        transferOptions.TransferMode = TransferMode.Binary; //The Transfer Mode - 
        //                                                                            //<em style="font-size: 9pt;">Automatic, Binary, or Ascii  
        //                        transferOptions.FilePermissions = null; //Permissions applied to remote files; 
        //                                                                //null for default permissions.  Can set user, 
        //                                                                //Group, or other Read/Write/Execute permissions. 
        //                        transferOptions.PreserveTimestamp = false; //Set last write time of 
        //                                                                   //destination file to that of source file - basically change the timestamp 
        //                                                                   //to match destination and source files.   
        //                        transferOptions.ResumeSupport.State = TransferResumeSupportState.Off;

        //                        TransferOperationResult transferResult;
        //                        //the parameter list is: local Path, Remote Path, Delete source file?, transfer Options  
        //                        transferResult = session.PutFiles(localPathtoUpload + filenameToUpload, RemotePath, false, transferOptions);
        //                        //Throw on any error 
        //                        transferResult.Check();
        //                        this.ltb_log.Items.Add(string.Format("{0} - Archivo enviado: {1}", DateTime.Now, fileInfo.Name));
        //                        this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //                    }
        //                    else
        //                    {
        //                        MessageBox.Show("Conexion error....", "FTP", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //                    }
        //                }
        //                //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
        //                ActualizaTransaccionEnviado(fileInfo.Name.Substring(7, 11));

        //                if (!File.Exists(str))
        //                {
        //                    File.Move(fileInfo.FullName, str);
        //                }
        //                else if (File.Exists(str))
        //                {
        //                    File.Delete(fileInfo.FullName);
        //                }
        //            }


        //            this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
        //            this.ltb_log.Items.Add("");
        //            this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //        }
        //        catch (Exception)
        //        {
        //            //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
        //            ActualizaTransaccionEnviado(_fName.Substring(7, 11));
        //            if (!File.Exists(str))
        //            {
        //                File.Move(strFullNme, str);
        //            }
        //            else if (File.Exists(str))
        //            {
        //                File.Delete(strFullNme);
        //            }

        //            this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
        //            this.ltb_log.Items.Add("");
        //            this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //            //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //    }
        //    private void ActualizaTransaccionEnviado(string _xReferenceNumber)
        //    {
        //        try
        //        {
        //            SqlCommand sqlCommand = new SqlCommand("SP_Actualiza_Transacciones_214_20", this._xConnString)
        //            {
        //                CommandType = CommandType.StoredProcedure
        //            };

        //            sqlCommand.Parameters.AddWithValue("@REFERENCENUMBER", _xReferenceNumber);
        //            this._xConnString.Open();
        //            sqlCommand.ExecuteNonQuery();
        //            this._xConnString.Close();
        //        }
        //        catch
        //        {
        //            this.ltb_log.Items.Add(string.Format("{0} - Problemas al cerrar el viaje.", DateTime.Now));
        //            this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //        }
        //    }
        //#endregion

        //#region Metodos_publicos
        //    public static Coordenates Convert2DegressMinutesSeconds(double position, bool isLong)
        //    {
        //        Coordenates ret = new Coordenates();

        //        // Negative: North
        //        // Positive: South
        //        // if isLong = true
        //        // Negative: East
        //        // Positive: West
        //        double absValue = Math.Abs(Math.Round(position * 1000000));
        //        int sign = Math.Sign(position);

        //        ret.Degress = (int)Math.Floor(absValue / 1000000);
        //        ret.Minutes = (int)Math.Floor(((absValue / 1000000) - Math.Floor(absValue / 1000000)) * 60);
        //        ret.Seconds = (Decimal)Math.Floor(((((absValue / 1000000) - Math.Floor(absValue / 1000000)) * 60) - Math.Floor(((absValue / 1000000) - Math.Floor(absValue / 1000000)) * 60)) * 100000) * 60 / 100000;

        //        if (isLong)
        //            if (sign > 0)
        //                ret.Geo = "W";
        //            else
        //                ret.Geo = "E";
        //        else
        //            if (sign > 0)
        //            ret.Geo = "N";
        //        else
        //            ret.Geo = "S";

        //        return ret;
        //    }
        //    public static Coordenates Convert2DegressMinutesSeconds(double position)
        //    {
        //        // By default, it is a Latitude (North / South)
        //        return Convert2DegressMinutesSeconds(position, false);
        //    }
        //#endregion

        //#region Estructura
        //    public struct Coordenates
        //    {
        //        public int Degress;
        //        public int Minutes;
        //        public decimal Seconds;
        //        public String Geo; // N,S,E,W
        //    }


        //private void GetTremec_Files()
        //{

        //    string _fName = "";

        //    string strFullNme = "";
        //    string FaltaVaor = "FaltaValor";



        //    int tb_EmbarqueDHL = 11;
        //    int tb_Orden = 2;
        //    int tb_IDOrigen = 0;
        //    int tb_RFCRemitente2 = 2;
        //    //SinValor
        //    int tb_Calle = 3;
        //    int tb_Municipio = 4;
        //    //SinValor
        //    int tb_Estado = 5;
        //    int tb_Pais = 6;
        //    int tb_CodigoPostal = 7;
        //    int tb_IDDestino = 0;
        //    int tb_RFCDestinatario2 = 9;
        //    //SinValor
        //    int tb_Calle2 = 10;
        //    int tb_Municipio3 = 11;
        //    int tb_Estado4 = 12;
        //    int tb_Pais5 = 13;
        //    int tb_CodigoPostal6 = 14;
        //    int tb_PesoNetoTotal = 15;
        //    int tb_NumTotalMercancias = 16;



        //    int tb_Descripción = 1;
        //    int tb_BienesTransp = 5;
        //    int tb_ClaveUnidad = 0;
        //    int tb_CveMaterialPeligroso = 10;
        //    int tb_FraccionArancelaria = 11;
        //    int tb_ValorMercancia = 11;


        //    int tb_Embalaje = 23;
        //    int tb_DescripEmbalaje = 24;
        //    int tb_UUIDComercioExt = 27;
        //    int tb_TotalKMRuta = 28;
        //    int InicioHeader = 16;
        //    int InicioTabla = 17;


        //    int[] ordenColm = new int[] { tb_EmbarqueDHL, tb_Orden, tb_IDOrigen, tb_RFCRemitente2, tb_Calle, tb_Municipio, tb_Estado, tb_Pais, tb_CodigoPostal, tb_IDDestino, tb_RFCDestinatario2, tb_Calle2, tb_Municipio3, tb_Estado4, tb_Pais5, tb_CodigoPostal6, tb_PesoNetoTotal, tb_NumTotalMercancias, tb_BienesTransp, tb_Descripción, tb_ClaveUnidad, tb_CveMaterialPeligroso, tb_Embalaje, tb_DescripEmbalaje, tb_ValorMercancia, tb_FraccionArancelaria, tb_UUIDComercioExt, tb_TotalKMRuta, InicioHeader, InicioTabla };
        //    DataTable errorTable = new DataTable();


        //    var directories = @"\\10.1.1.30\e$\Attachments\sftr-tremec";
        //    bool direxists = System.IO.Directory.Exists(directories);



        //    try
        //    {
        //        // Setup session options

        //        if (direxists)
        //        {

        //            //var directories = Directory.GetDirectories(@"\\10.1.1.30\e$\Attachments\sftr-dhlsupply");
        //            FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
        //            for (int j = 0; j < (int)files.Length; j++)
        //            {
        //                FileInfo fileInfo = files[j];
        //                if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
        //                {

        //                    String filenameToUpload = fileInfo.Name;
        //                    _fName = fileInfo.Name;
        //                    strFullNme = fileInfo.FullName;

        //                    FileStream fileStream = null;
        //                    FileInfo datos = new FileInfo(fileInfo.DirectoryName + "\\" + fileInfo.Name);
        //                    ExcelPackage LecExcel = new ExcelPackage(datos);
        //                    try
        //                    {
        //                        fileStream = datos.Open(FileMode.Open, FileAccess.ReadWrite);
        //                    }
        //                    catch (NullReferenceException)
        //                    {
        //                        MessageBox.Show("El archivo se encuentra abierto por otro proceso.");
        //                        return;
        //                    }
        //                    catch (IOException)
        //                    {
        //                        MessageBox.Show("El archivo se encuentra abierto por otro proceso.");
        //                        return;
        //                    }
        //                    LecExcel.Load(fileStream);

        //                    ExcelWorksheet worksheet = LecExcel.Workbook.Worksheets[2];
        //                    fileStream.Close();

        //                    if (worksheet.Dimension == null)
        //                    {
        //                        return;
        //                    }

        //                    //var dateString = "12-12-2021 12:00:00";
        //                    //DateTime date1 = DateTime.Parse(dateString, System.Globalization.CultureInfo.CurrentCulture);

        //                    DataTable excelDT = new DataTable();
        //                    excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 2, InicioHeader, InicioTabla);
        //                    errorTable = excelDT;
        //                    if (excelDT.Rows.Count == 0)
        //                    {
        //                        ErrorMessageTxt += String.Format("El archivo no registro ninguna tabla, es posible que se trate de otro formato o el inicio de la tabla no sea correcta   {0}", "<br />");
        //                        EnviarEmail(directories, fileInfo.Name, ErrorMessageTxt);
        //                    }
        //                    //El primer numero determina cuando comienza el header el segundo numero es donde comienza la tabla
        //                    //6 = header, 8 = comienzo de la tabla

        //                    String RSRemitent = worksheet.Cells[2, 2].Value?.ToString();
        //                    String RFCRemitent = worksheet.Cells[2, 8].Value?.ToString();
        //                    String CalleRemitent = worksheet.Cells[5, 2].Value?.ToString();
        //                    String MunicipioRemitent = worksheet.Cells[5, 3].Value?.ToString();
        //                    String EstadoRemitent = worksheet.Cells[5, 4].Value?.ToString();
        //                    String PaisRemitent = worksheet.Cells[5, 5].Value?.ToString();

        //                    String CodigoPRemitent = worksheet.Cells[5, 6].Value?.ToString();

        //                    String CalleEntrega = worksheet.Cells[12, 2].Value?.ToString();
        //                    String MunicipioEntrega = worksheet.Cells[11, 3].Value?.ToString();
        //                    String EstadoEntrega = worksheet.Cells[11, 4].Value?.ToString();
        //                    String PaisEntrega = worksheet.Cells[11, 5].Value?.ToString();
        //                    String CodigoPEntrega = worksheet.Cells[11, 6].Value?.ToString();

        //                    String PesoBrutoT = worksheet.Cells[24, 5].Value?.ToString();
        //                    String NumeroTMercancia = worksheet.Cells[24, 7].Value?.ToString();


        //                    List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

        //                    foreach (DataRow e in excelDT.Rows)
        //                    {

        //                        if (PesoBrutoT != null)
        //                        {

        //                            EntintDHLFields ent = new EntintDHLFields();

        //                            ent.ReferenciaDelServicio = FaltaVaor;
        //                            ent.RSdelRemitente = RSRemitent;
        //                            ent.RFCdelRemitente = RFCRemitent;
        //                            ent.Supplier = FaltaVaor;
        //                            ent.Calle = CalleRemitent;
        //                            ent.Municipio = MunicipioRemitent;
        //                            ent.Estado = EstadoRemitent;
        //                            ent.Pais = PaisRemitent;
        //                            ent.CP = CodigoPRemitent;
        //                            ent.RSdelDestinatario = FaltaVaor;
        //                            ent.RFCDestinatario = FaltaVaor;
        //                            ent.Calle2 = CalleEntrega;
        //                            ent.Municipio2 = CalleEntrega;
        //                            ent.Estado2 = EstadoEntrega;
        //                            ent.Pais2 = PaisEntrega;
        //                            ent.CP2 = CodigoPEntrega;
        //                            ent.PesoNeto = PesoBrutoT;
        //                            ent.NumeroTotalMercancias = NumeroTMercancia;
        //                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString();
        //                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString();
        //                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString();
        //                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
        //                            ent.ValorDeLaMercancia = FaltaVaor;
        //                            ent.TipoDeMoneda = FaltaVaor;

        //                            if (ent.ReferenciaDelServicio != "")
        //                                eDhlList.Add(ent);

        //                        }
        //                        else
        //                        {
        //                            //Mandar email
        //                            ErrorMessageTxt += String.Format("Un campo requerido no esta en la tabla{0}", "<br />");
        //                            EnviarEmail(directories, fileInfo.Name, DetectarErrorFormato(excelDT, ordenColm));
        //                            break;
        //                        }

        //                    }
        //                    if (eDhlList.Any())
        //                    {
        //                        List<EntintDHLFields> ClientList = new List<EntintDHLFields>();
        //                        ClientList = VerificaMercanciasClientes(eDhlList);
        //                        bool isEmpt = !ClientList.Any();
        //                        if (isEmpt) { }
        //                        else
        //                        {
        //                            InsertaMercanciaClientesList(ClientList, directories, fileInfo.Name);
        //                        }
        //                    }
        //                }
        //            }

        //            CortarDocumentos(directories);

        //        }
        //        this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
        //        this.ltb_log.Items.Add("");
        //        //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //    }
        //    catch (Exception ex)
        //    {
        //        string erro = ex.ToString();
        //        //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
        //        ErrorMessageTxt += String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
        //        EnviarEmail(directories, _fName, ErrorMessageTxt);

        //        CortarDocumentosPError(directories, _fName);
        //        //ActualizaTransaccionEnviado(_fName.Substring(7, 11));

        //        this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion Tremec...", DateTime.Now));
        //        this.ltb_log.Items.Add("");
        //        //this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
        //        //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //}
        #endregion
    }

    //private void GetAmazon_Files()
    //{

    //    string _fName = "";
    //    string str = "";
    //    string strFullNme = "";


    //    int tb_EmbarqueDHL = 0;
    //    int tb_Orden = 2;
    //    int tb_IDOrigen = 0;
    //    int tb_RFCRemitente2 = 2;
    //    int tb_RSRemitente2 = 1;
    //    int tb_RSDest = 8;
    //    //SinValor
    //    int tb_Calle = 3;
    //    int tb_Municipio = 4;
    //    //SinValor
    //    int tb_Estado = 5;
    //    int tb_Pais = 6;
    //    int tb_CodigoPostal = 7;
    //    int tb_IDDestino = 0;
    //    int tb_RFCDestinatario2 = 9;
    //    //SinValor
    //    int tb_Calle2 = 10;
    //    int tb_Municipio3 = 11;
    //    int tb_Estado4 = 12;
    //    int tb_Pais5 = 13;
    //    int tb_CodigoPostal6 = 14;
    //    int tb_PesoNetoTotal = 15;
    //    int tb_NumTotalMercancias = 16;
    //    int tb_BienesTransp = 17;
    //    int tb_Descripción = 18;
    //    int tb_ClaveUnidad = 23;
    //    int tb_CveMaterialPeligroso = 20;
    //    int tb_Embalaje = 23;
    //    int tb_DescripEmbalaje = 24;
    //    int tb_ValorMercancia = 21;
    //    int tb_FraccionArancelaria = 22;
    //    int tb_UUIDComercioExt = 27;
    //    int tb_TotalKMRuta = 28;


    //    var directories = @"\\10.1.1.30\FTProot\sftr-amazon";


    //    
    //    try
    //    {
    //        // Setup session options
    //        SessionOptions sessionOptions = new SessionOptions
    //        {
    //            Protocol = Protocol.Ftp,
    //            FtpSecure = FtpSecure.Explicit,
    //            FtpMode = FtpMode.Passive,
    //            PortNumber = Constantes.PuertoSFTP,
    //            HostName = Constantes.ServidorSFTP, //this.txtHostName.Text,
    //            UserName = Constantes.UsuarioSFTP, //this.txtUserName.Text,
    //            Password = Constantes.PassWordSFTP, //this.txtPwd.Text,
    //            GiveUpSecurityAndAcceptAnySshHostKey = false,
    //            GiveUpSecurityAndAcceptAnyTlsHostCertificate = true
    //        };
    //        session.Open(sessionOptions);

    //        //var directories = Directory.GetDirectories(@"\\10.1.1.30\FTProot\sftr-dhlsupply");
    //        FileInfo[] files = (new DirectoryInfo(directories)).GetFiles();
    //        for (int j = 0; j < (int)files.Length; j++)
    //        {
    //            FileInfo fileInfo = files[j];
    //            if ((fileInfo.Name.Contains(".xlsx") ? true : fileInfo.Name.Contains(".xlsx")))
    //            {

    //                String filenameToUpload = fileInfo.Name;
    //                _fName = fileInfo.Name;
    //                strFullNme = fileInfo.FullName;

    //                if (session.Opened)
    //                {

    //                    DataTable excelDT = new DataTable();
    //                    excelDT = ProcesoExcel(excelDT, fileInfo.DirectoryName + "\\" + fileInfo.Name, 1, 1, 2);
    //                    //El primer numero determina cuando comienzala hoja, despues  el header el tercer numero es donde comienza la tabla
    //                    //1 = hoja, 6 = header, 8 = comienzo de la tabla

    //                    List<EntintDHLFields> eDhlList = new List<EntintDHLFields>();

    //                    foreach (DataRow e in excelDT.Rows)
    //                    {

    //                        if ((e.ItemArray[tb_EmbarqueDHL].ToString() != null) && (e.ItemArray[tb_EmbarqueDHL].ToString() != "")
    //                            && (e.ItemArray[tb_Municipio].ToString() != null) && (e.ItemArray[tb_Municipio].ToString() != "")
    //                            && (e.ItemArray[tb_Municipio3].ToString() != null) && (e.ItemArray[tb_Municipio3].ToString() != "")
    //                                    && (excelDT.Columns["Estado"].Ordinal == tb_Estado))
    //                        {

    //                            EntintDHLFields ent = new EntintDHLFields();

    //                            ent.ReferenciaDelServicio = e.ItemArray[tb_EmbarqueDHL].ToString();
    //                            ent.RSdelRemitente = e.ItemArray[tb_RSRemitente2].ToString();
    //                            ent.RFCdelRemitente = e.ItemArray[tb_RFCRemitente2].ToString();
    //                            ent.Supplier = e.ItemArray[tb_IDOrigen].ToString();
    //                            ent.Calle = e.ItemArray[tb_Calle].ToString();
    //                            ent.Municipio = e.ItemArray[tb_Municipio].ToString();
    //                            ent.Estado = e.ItemArray[tb_Estado].ToString();
    //                            ent.Pais = e.ItemArray[tb_Pais].ToString();
    //                            ent.CP = e.ItemArray[tb_CodigoPostal].ToString().Trim();
    //                            ent.RSdelDestinatario = e.ItemArray[tb_RSDest].ToString();
    //                            ent.RFCDestinatario = e.ItemArray[tb_RFCDestinatario2].ToString();
    //                            ent.Calle2 = e.ItemArray[tb_Calle2].ToString();
    //                            ent.Municipio2 = e.ItemArray[tb_Municipio3].ToString();
    //                            ent.Estado2 = e.ItemArray[tb_Estado4].ToString();
    //                            ent.Pais2 = e.ItemArray[tb_Pais5].ToString();
    //                            ent.CP2 = e.ItemArray[tb_CodigoPostal6].ToString().Trim();
    //                            ent.PesoNeto = e.ItemArray[tb_PesoNetoTotal].ToString();
    //                            ent.NumeroTotalMercancias = e.ItemArray[tb_NumTotalMercancias].ToString();
    //                            ent.ClaveDelBienTransportado = e.ItemArray[tb_BienesTransp].ToString().Trim(new Char[] { '\'' });
    //                            ent.ClaveUnidadDeMedida = e.ItemArray[tb_ClaveUnidad].ToString();
    //                            ent.DescripcionDelBienTransportado = e.ItemArray[tb_Descripción].ToString().Replace('|', ' ').Replace(',', ' ').Replace('°', ' ').Replace('-', ' ')
    //                                .Replace('\\', ' ').Replace('/', ' ').Replace('\'', ' ').Replace('’', ' ').Replace('"', ' ');
    //                            ent.MaterialPeligroso = e.ItemArray[tb_CveMaterialPeligroso].ToString();
    //                            ent.ValorDeLaMercancia = e.ItemArray[tb_ValorMercancia].ToString();
    //                            ent.TipoDeMoneda = e.ItemArray[tb_FraccionArancelaria].ToString();

    //                            if (ent.ReferenciaDelServicio != "")
    //                                eDhlList.Add(ent);

    //                        }
    //                        else
    //                        {
    //                            //Mandar email
    //                            EnviarEmail(directories, fileInfo.Name);
    //                        }

    //                    }
    //                    if (eDhlList.Any())
    //                        InsertaMercanciaClientesList(eDhlList, directories, fileInfo.Name);
    //                }
    //                else
    //                {
    //                    MessageBox.Show("Conexion error....", "FTP", MessageBoxButtons.OK, MessageBoxIcon.Error);
    //                }
    //            }
    //        }

    //        CortarDocumentos(directories);

   
    //        this.ltb_log.Items.Add(string.Format("{0} - Proceso Completado...", DateTime.Now));
    //        this.ltb_log.Items.Add("");
    //        this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
    //    }
    //    catch (Exception ex)
    //    {
    //        string erro = ex.ToString();
    //        //ACTUALIZAMOS EL ESTATUS DE LA TRANSACCION A ENVIADO.
    //        EnviarEmail(directories, _fName);
    //        CortarDocumentosPError(directories, _fName);
    //        ActualizaTransaccionEnviado(_fName.Substring(7, 11));
    //        if (!File.Exists(str))
    //        {
    //            File.Move(strFullNme, str);
    //        }
    //        else if (File.Exists(str))
    //        {
    //            File.Delete(strFullNme);
    //        }
   
    //        this.ltb_log.Items.Add(string.Format("{0} - Error en la operacion...", DateTime.Now));
    //        this.ltb_log.Items.Add("");
    //        this.ltb_log.SelectedIndex = this.ltb_log.Items.Count - 1;
    //        //MessageBox.Show(ex.Message, "FTP error de componente", MessageBoxButtons.OK, MessageBoxIcon.Error);
    //    }
    //}


    //ent.ReferenciaDelServicio = e["Embarque DHL"].ToString();
    //ent.RSdelRemitente = e["IDOrigen"].ToString();
    //ent.RFCdelRemitente = e["RFCRemitente2"].ToString();
    //ent.Supplier = e["IDOrigen"].ToString();
    //ent.Calle = e["Calle"].ToString();
    //ent.Municipio = e["Municipio"].ToString();
    //ent.Estado = e["Estado"].ToString();
    //ent.Pais = e["Pais"].ToString();
    //ent.CP = e["CodigoPostal"].ToString();
    //ent.RSdelDestinatario = e["IDDestino"].ToString();
    //ent.RFCDestinatario = e["RFCDestinatario2"].ToString();
    //ent.Calle2 = e["Calle2"].ToString();
    //ent.Municipio2 = e["Municipio3"].ToString();
    //ent.Estado2 = e["Estado4"].ToString();
    //ent.Pais2 = e["Pais5"].ToString();
    //ent.CP2 = e["CodigoPostal6"].ToString();
    //ent.PesoNeto = e["PesoNetoTotal"].ToString();
    //ent.NumeroTotalMercancias = e["NumTotalMercancias"].ToString();
    //ent.ClaveDelBienTransportado = e["BienesTransp"].ToString();
    //ent.ClaveUnidadDeMedida = e["ClaveUnidad"].ToString();
    //ent.DescripcionDelBienTransportado = e["Descripción"].ToString();
    //ent.MaterialPeligroso = e["CveMaterialPeligroso"].ToString();
    //ent.ValorDeLaMercancia = e["ValorMercancia"].ToString();
    //ent.TipoDeMoneda = e["FraccionArancelaria"].ToString();


}