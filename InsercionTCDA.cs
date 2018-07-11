using MySql.Data.MySqlClient;
using System;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace VolcadoDeAdjuntos {
    class InsercionTCDA {
        //Creamos la cadena de conexion a la base de datos.
        static String cred = "SERVER=" + ConfigurationManager.AppSettings.Get("ipbd").ToString() + ";DATABASE=" + ConfigurationManager.AppSettings.Get("database").ToString() + ";UID=" + ConfigurationManager.AppSettings.Get("userbd").ToString() + ";PWD=" + ConfigurationManager.AppSettings.Get("passbd").ToString() + ";";

        public static void InsercionData(String[,] cellData, int registros, DateTime fe) {
            MySqlConnection conn = new MySqlConnection(cred);
            MySqlCommand cmd;
            //Abrimos la conexion a la base de datos utilizando la cadena de conexion.
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
                        //Si el campo i, 7 que es una fecha es null, le asignamos el formato de fecha por defecto
                        if (cellData[i, 7] == null || cellData[i, 7] == "") {
                            cellData[i, 7] = DefaultDate();
                        } else {
                            //Si el campo no es null, cambiamos su formato para que se adapte a lo que se quiere.
                            cellData[i, 7] = DateConverter(cellData, i, 7);
                        }

                        if (cellData[i, 8] == null || cellData[i, 8] == "") {
                            cellData[i, 8] = DefaultDate();
                        } else {
                            cellData[i, 8] = DateConverter(cellData, i, 8);
                        }

                        if (cellData[i, 10] == null || cellData[i, 10] == "") {
                            cellData[i, 10] = DefaultDate();
                        } else {
                            cellData[i, 10] = DateConverter(cellData, i, 10);
                        }

                        if (cellData[i, 11] == null || cellData[i, 11] == "") {
                            cellData[i, 11] = DefaultDate();
                        } else {
                            cellData[i, 11] = DateConverter(cellData, i, 11);
                        }

                        if (cellData[i, 18] == null || cellData[i, 18] == "") {
                            cellData[i, 18] = "0";
                        }

                        //Extraemos los caracteres ínvalidos del campo OperationComment.
                        String normalizado = cellData[i, 3].Normalize(System.Text.NormalizationForm.FormD);
                        Regex reg = new Regex("[']");
                        cellData[i, 3] = reg.Replace(normalizado, "");

                        try {
                            cmd.CommandText = "INSERT INTO tickets_cerrados_da (`Lastmodified`, `Ticket ID`, `Third Party reference`, `Identifier 1`, `Company name`, `Initiator - User name`, `Closure user name`, `Current action`, `tickets.creation date`, `closure date`, `Transfer date`, `Restoration date (UTC)`, `Last Resolution date (UTC)`, `Ticket dur. BH 8am - 8pm Mon to Fri`, `Time To Repair`, `Time To Resolv`, `Ticket Type`, `Problem detail`, `Short label`, `Recipient - Group ID`, `Closure group ID`, `Initiating group ID`, `Initial nature`, `Request date`, `Indentifier 3`, `Contractual duration (mm)`, `Cause Label`) VALUES('" + dat + "', '" + cellData[i, 0] + "','" + cellData[i, 1] + "','" + cellData[i, 2] + "','" + cellData[i, 3] + "','" + cellData[i, 4] + "','" + cellData[i, 5] + "','" + cellData[i, 6] + "','" + cellData[i, 7] + "','" + cellData[i, 8] + "','" + cellData[i, 9] + "','" + cellData[i, 10] + "','" + cellData[i, 11] + "'," + cellData[i, 12] + "," + cellData[i, 13] + "," + cellData[i, 14] + ",'" + cellData[i, 15] + "','" + cellData[i, 16] + "', '" + cellData[i, 17] + "', '" + cellData[i, 18] + "', " + cellData[i, 19] + ", " + cellData[i, 20] + ", '" + cellData[i, 21] + "','" + cellData[i, 22] + "', '" + cellData[i, 23] + "', " + cellData[i, 24] + ", '" + cellData[i, 25] + "')";
                            cmd.ExecuteNonQuery();
                        }catch(Exception e) {
                            LOGS.Log("InsercionTCDA 1erSubCatch --> " + e.StackTrace + "\n");
                            //MessageBox.Show("" + e);
                            //cmd.CommandText = "INSERT INTO tickets_cerrados_da (`Lastmodified`, `Ticket ID`, `Third Party reference`, `Identifier 1`, `Company name`, `Initiator - User name`, `Closure user name`, `Current action`, `tickets.creation date`, `closure date`, `Transfer date`, `Restoration date (UTC)`, `Last Resolution date (UTC)`, `Ticket dur. BH 8am - 8pm Mon to Fri`, `Time To Repair`, `Time To Resolv`, `Ticket Type`, `Problem detail`, `Short label`, `Recipient - Group ID`, `Closure group ID`, `Initiating group ID`, `Initial nature`, `Request date`, `Contractual duration (mm)`) VALUES('" + dat + "', '" + cellData[i, 0] + "','" + cellData[i, 1] + "','" + cellData[i, 2] + "','" + cellData[i, 3] + "','" + cellData[i, 4] + "','" + cellData[i, 5] + "','" + cellData[i, 6] + "','" + cellData[i, 7] + "','" + cellData[i, 8] + "','" + cellData[i, 9] + "','" + cellData[i, 10] + "','" + cellData[i, 11] + "','" + cellData[i, 12] + "'," + cellData[i, 13] + "," + cellData[i, 14] + ",'" + cellData[i, 15] + "','" + cellData[i, 16] + "', '" + cellData[i, 17] + "', " + cellData[i, 18] + ", " + cellData[i, 19] + ", " + cellData[i, 20] + ", '" + cellData[i, 21] + "','" + cellData[i, 22] + "', " + cellData[i, 23] + ")";
                            //cmd.ExecuteNonQuery();
                        }
                    } catch (Exception e) {
                        LOGS.Log("InsercionTCDA 1er Catch --> " + e.StackTrace + "\n");
                        if (conn.State == System.Data.ConnectionState.Open) {
                            conn.Close();
                        }
                    }
                }
                LOGS.Log("TCDA --> INSERCION --> OK");
            } catch (Exception ex) {
                LOGS.Log("InsercionTCDA (InsercionData) --> " + ex + "\n");
                throw;
            } finally {
                if (conn.State == System.Data.ConnectionState.Open) {
                    conn.Close();
                }
            }
        }

        //Método que cambia el formato de la fecha del campo fecha del excel abierto.
        public static String DateConverter(String[,] Array, int indiceFor, int indiceCambio) {
            DateTime dataValue;
            String formatDates = null;
            try {
                dataValue = DateTime.Parse(Array[indiceFor, indiceCambio]);
                formatDates = dataValue.ToString("yyyy/MM/dd HH:mm:ss");
            } catch (Exception date) {
                LOGS.Log("InsercionTCDA (DateConverter) --> " + date.StackTrace + "\n");
            }
            return formatDates;
        }

        //Devolvemos una fecha por defecto para aquellos campos nulos que requieran una fecha.
        public static String DefaultDate() {
            DateTime dataValue = new DateTime(1971, 01, 01, 00, 00, 00); ;
            String formatDates = null;
            try {
                formatDates = dataValue.ToString("yyyy/MM/dd HH:mm:ss");
            } catch (Exception def) {
                LOGS.Log("InsercionTCDA (DefaultData) --> " + def.StackTrace + "\n");
            }
            return formatDates;
        }

    }
}