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
            this.CargarConexion(Constantes._xServidorBD, Constantes._xNombreBD, Constantes._xTrusted_Connection);
        }


        public void CargarConexion(string _xServerName, string _xNombreBD, string _xTrustedConnection)
        {
            //server = DESKTOP - FP59UDN\\SQLEXPRESS; database = x_FTR; Trusted_Connection = true
            try
            {
                this._xConnString.ConnectionString = string.Format("server = {0}; database = {1}; Trusted_Connection = {2};", new object[] { _xServerName, _xNombreBD, _xTrustedConnection });
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
            //this.UploadFiles();
            try
            {
                DataTable dt = new DataTable();
                string sentencia = string.Format("SELECT [idCliente],[nombre], [fechaNacimiento] FROM [Recursos].[dbo].[baseClientesHackaton2022] ");
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
            }
            catch
            {
                this.ConsoladeSalida = "Problemas al obtener catalogos los archivos.";
            }
            
        }
    }
}
