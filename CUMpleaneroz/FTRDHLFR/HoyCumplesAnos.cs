using CUMpleanero;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FTRDHLFR
{
    public class HoyCumplesAnos
    {
        private SqlConnection _xConnString = new SqlConnection();
        public string ConsoladeSalida = "";
        public HoyCumplesAnos()
        {
            this.CargarConexion(Constantes._xServidorBD, Constantes._xNombreBD, Constantes._xUsuarioBD, Constantes._xPassWordBD);
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

        public void StartCheck()
        {
            this.ConsoladeSalida = "";
            ChecarSiCumplesAnosHoy();
        }

        public void ChecarSiCumplesAnosHoy()
        {
            string FinalCollage = "";
            //this.UploadFiles();
            try
            {
                DataTable dt = new DataTable();
                string sentencia = string.Format("SELECT TOP 2000 [empl_id],[name], [fec_naci],[status] FROM [x_FTR].[dbo].[em] WHERE status <> 'TER' AND MONTH([fec_naci]) = DATEPART(MONTH,getdate())  ORDER BY DAY([fec_naci]) ASC");
                this._xConnString.Open();

                SqlDataAdapter dataAda = new SqlDataAdapter(sentencia, this._xConnString);
                dataAda.Fill(dt);
                this._xConnString.Close();

                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        //StellantCopy.RemoveAt(i);
                    }
                }

                this.ConsoladeSalida = "La lista se pude leer, creo";
                //dt = OrdenarCumpleañosPorFecha(dt);
                FinalCollage = HacerStringParaHtmlMail(dt);
                this.ConsoladeSalida = FinalCollage;
                EnviarEmailCumple(FinalCollage);
            }
            catch
            {
                this.ConsoladeSalida = "Problemas al obtener catalogos los archivos.";
            }
            
        }

        #region privateRegion
        //Problema resuelto con el order bu date de sql pero esta bueno para ejercicio de orden de listas 
        //Problema resuelto con el order bu date de sql pero esta bueno para ejercicio de orden de listas 
        //Problema resuelto con el order bu date de sql pero esta bueno para ejercicio de orden de listas 
        public DataTable OrdenarCumpleañosPorFecha(DataTable dtTable)
        {
            DataTable dataTable = new DataTable();


            //Obtenemos la lista para ordenar
            //Obtenemos la lista para ordenar
            List<CumpleaneroEntity> OrdenList = new List<CumpleaneroEntity>();
            foreach (DataRow e in dtTable.Rows)
            {
                //Obtenemos el cumpleaños del empleado en la string datehoy
                string datehoy = e.ItemArray[2].ToString();
                //Obtenemos el cumpleaños del empleado en la string datehoy
                var parsedDate = DateTime.Parse(e.ItemArray[2].ToString());


                CumpleaneroEntity ent = new CumpleaneroEntity();

                //Asignando valores a la entidad
                //Asignando valores a la entidad
                ent.IdEmployee = e.ItemArray[0].ToString();
                ent.Nombre = e.ItemArray[1].ToString();
                ent.fechaNacimiento = parsedDate.Day;
                //Asignando valores a la entidad
                //Asignando valores a la entidad

                OrdenList.Add(ent);
            }
            //Obtenemos la lista para ordenar
            //Obtenemos la lista para ordenar


            //Codigo para ordenar la lista
            //Codigo para ordenar la lista
            //Codigo para ordenar la lista
            //Codigo para ordenar la lista


            return dtTable;
        }
        #endregion

        public string HacerStringParaHtmlMail(DataTable dtTable)
        {
            string CollageMensajes = "";
            // 0 = Idempl
            // 1 = nombre
            // 2 = fecha
            foreach (DataRow e in dtTable.Rows)
            {
                //String.Format("El archivo no pudo leerse Correctamente{0}", "<br />");
                //Acomoda El nombre 
                if ((e.ItemArray[1].ToString() != null) && (e.ItemArray[1].ToString() != ""))
                {
                    CollageMensajes += String.Format("<tr> <td>{0}  </td>", e.ItemArray[1].ToString()); 
                }
                //Acomoda La fecha
                if ((e.ItemArray[2].ToString() != null) && (e.ItemArray[2].ToString() != ""))
                {
                    //DateTime DT1 = new DateTime(datehoy);
                    //var cultureInfo = new CultureInfo("es-ES");
                    string datehoy = e.ItemArray[2].ToString();
                    var parsedDate = DateTime.Parse(datehoy);
                    CollageMensajes += String.Format("<td>  {0} - {1}</td> {2}", parsedDate.Day, parsedDate.ToString("MMMM"), "<tr/>");
                }
            }
            return CollageMensajes;
        }
        public void EnviarEmailCumple(string mensaje)
        {
            string destEmail = "";
            string msgMail1 = "Felicidades compañeros.";
            string msgMail2 = mensaje;
            string msgMail3 = "Contacto javega@ftr.com.mx ";
            string asunto = "Cumpleaños";


            //destEmail = "desaftr02@ftr.com.mx";
            destEmail = "eanavia@ftr.com.mx";
            //Cambiar a true            //Cambiar a true            //Cambiar a true
            SendMail smail = new SendMail(true);
            smail.Notificar(destEmail, msgMail1, msgMail2, msgMail3, asunto);
        }
    }
}
