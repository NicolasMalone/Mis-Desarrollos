using MySql.Data.MySqlClient;
using System;
using System.Configuration;
using System.Windows.Forms;

namespace VolcadoDeAdjuntos {
    class InsercionBO {

        //Se crea la cadena de conexion.
        static String cred = "SERVER=" + ConfigurationManager.AppSettings.Get("ipbd").ToString() + ";DATABASE=" + ConfigurationManager.AppSettings.Get("database").ToString() + ";UID=" + ConfigurationManager.AppSettings.Get("userbd").ToString() + ";PWD=" + ConfigurationManager.AppSettings.Get("passbd").ToString() + ";";

        public static void InsercionData(String[,] cellData, int registros, DateTime fe) {
            MySqlConnection conn = new MySqlConnection(cred);
            MySqlCommand cmd;
            //Se abre la conexion a la base de datos
            conn.Open();

            try {

                String dat = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                //fe = fe.AddHours(Double.Parse(DateTime.Now.ToString("HH")));
                //fe = fe.AddMinutes(Double.Parse(DateTime.Now.ToString("mm")));
                //fe = fe.AddSeconds(Double.Parse(DateTime.Now.ToString("ss")));
                //String dat = fe.ToString("yyyy/MM/dd HH:mm:ss");

                cmd = conn.CreateCommand();

                for (int i = 0; i < registros - 5; i++) {
                    try {
                        //longitudes de campos en BBDD (Exceptuando Fechas en formato TimeStamp).
                        cellData[i, 16] = DateConverter(cellData, i, 16);

                        if (cellData[i, 10].Length > 500) {
                            cellData[i, 10] = cellData[i, 10].Substring(0, 499);
                        }                        

                        cmd.CommandText = "INSERT INTO backlog_oceane (`Lastmodified`, `Ticket ID`, `Owner - Group ID`, `Ticket type`, `Identifier 1`, `Company name`, `Current action`, `Third party reference`, `Short label`, `Final nature`, `Problem family`, `Problem datail`, `Cause label`, `Initiator - User name`, `Closure user name`,  `Initiating group ID`, `Closure group ID`, `Creation date`, `Owner - Group ID II`) VALUES('" + dat + "', '" + cellData[i, 0] + "','" + cellData[i, 1] + "','" + cellData[i, 2] + "','" + cellData[i, 3] + "','" + cellData[i, 4] + "','" + cellData[i, 5] + "','" + cellData[i, 6] + "','" + cellData[i, 7] + "','" + cellData[i, 8] + "','" + cellData[i, 9] + "','" + cellData[i, 10] + "','" + cellData[i, 11] + "','" + cellData[i, 12] + "','" + cellData[i, 13] + "','" + cellData[i, 14] + "','" + cellData[i, 15] + "','" + cellData[i, 16] + "', '" + cellData[i, 18] + "')";
                        cmd.ExecuteNonQuery();
                    } catch (Exception e) {
                        LOGS.Log("InsercionBO 1erSubCatch --> " + e + "\n");
                        //MessageBox.Show("" + e);
                        if (conn.State == System.Data.ConnectionState.Open) {
                            conn.Close();
                        }
                    }
                }
                conn.Close();
            } catch (Exception ex) {
                LOGS.Log("InsercionBO 1erCatch --> " + ex.StackTrace + "\n");
                throw;
            } finally {
                if (conn.State == System.Data.ConnectionState.Open) {
                    conn.Close();
                }
            }
        }



        public static String DateConverter(String[,] Array, int indiceFor, int indiceCambio) {
            DateTime dataValue;
            String formatDates = null;

            try {
                dataValue = DateTime.Parse(Array[indiceFor, indiceCambio]);
                formatDates = dataValue.ToString("yyyy/MM/dd HH:mm:ss");
            } catch (Exception e) {
                LOGS.Log("InsercionBO (DateConverter) --> " + e.StackTrace + "\n");
            }
            return formatDates;
        }

        ////Devolvemos una fecha por defecto para aquellos campos nulos que requieran una fecha.
        //public static String DefaultDate() {
        //    DateTime dataValue;
        //    String formatDates;

        //    dataValue = DateTime.MinValue;
        //    formatDates = dataValue.ToString("yyyy/MM/dd HH:mm:ss");

        //    return formatDates;
        //}

    }
}