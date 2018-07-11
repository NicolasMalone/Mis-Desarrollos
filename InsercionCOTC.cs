using MySql.Data.MySqlClient;
using System;
using System.Configuration;
using System.Text.RegularExpressions;

namespace VolcadoDeAdjuntos {
    class InsercionCOTC {
        //Se crea la cadena de conexion a la base de datos.
        static String cred = "SERVER=" + ConfigurationManager.AppSettings.Get("ipbd").ToString() + ";DATABASE=" + ConfigurationManager.AppSettings.Get("database").ToString() + ";UID=" + ConfigurationManager.AppSettings.Get("userbd").ToString() + ";PWD=" + ConfigurationManager.AppSettings.Get("passbd").ToString() + ";";

        //Metodo que recibe el array de datos de los excels y el numero de registros o filas de datos del adjunto.
        public static void InsercionData(String[,] cellData, int registros, DateTime fe) {
            MySqlConnection conn = new MySqlConnection(cred);
            MySqlCommand cmd;
            //Se abre la conexion a la base de datos.
            

            try {
                String dat = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");

                //fe = fe.AddHours(Double.Parse(DateTime.Now.ToString("HH")));
                //fe = fe.AddMinutes(Double.Parse(DateTime.Now.ToString("mm")));
                //fe = fe.AddSeconds(Double.Parse(DateTime.Now.ToString("ss")));
                //String dat = fe.ToString("yyyy/MM/dd HH:mm:ss");

                cmd = conn.CreateCommand();
                //registros - 5 para evitar que los datos de las primeras 5 filas sean leidos innecesariamente
                for (int i = 0; i < registros - 5; i++) {
                    try {
                        conn.Open();
                        //Sobreescibimos las fechas de los campos.
                        cellData[i, 3] = DateConverter(cellData, i, 3);
                        cellData[i, 8] = DateConverter(cellData, i, 8);
                        cellData[i, 9] = DateConverter(cellData, i, 9);
                        //Extraemos los caracteres ínvalidos del campo OperationComment.
                        String normalizado = cellData[i, 4].Normalize(System.Text.NormalizationForm.FormD);
                        Regex reg = new Regex("[']");
                        cellData[i, 4] = reg.Replace(normalizado, "");
                        if (cellData[i, 2] == null || cellData[i, 2] == "") {
                            cellData[i, 2] = "0";
                        }

                        //Introducimos los datos del excel adjunto en la base de datos.
                        cmd.CommandText = "INSERT INTO comentario_oceane_tc (LastModified, TicketID, CommentID, CommentNumber, CommentDate, OperationComment, UserName, CommentType, CurrentAction, CreationDate, ClosureDate, TicketDuration, TicketType, ProblemDetail, ShortLabel, ClosureGroupID, OwnerGroupID) VALUES('" + dat + "', '" + cellData[i, 0] + "','" + cellData[i, 1] + "'," + cellData[i, 2] + ",'" + cellData[i, 3] + "','" + cellData[i, 4].ToString() + "','" + cellData[i, 5] + "','" + cellData[i, 6] + "','" + cellData[i, 7] + "','" + cellData[i, 8] + "','" + cellData[i, 9] + "'," + cellData[i, 10] + ",'" + cellData[i, 11] + "','" + cellData[i, 12] + "','" + cellData[i, 13] + "','" + cellData[i, 14] + "','" + cellData[i, 15] + "')";
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    } catch (FormatException) {
                        cellData[i, 3] = DateConverter(cellData, i, 8);
                    } catch (Exception e) {
                        LOGS.Log("InsercionCOTC 1erSubCatch --> " + e.StackTrace + "\n");
                    } finally {
                        if (conn.State == System.Data.ConnectionState.Open) {
                            conn.Close();
                        }
                    }
                }
                LOGS.Log("COTC --> INSERCION --> OK");
            } catch (Exception) {
                throw;
            }
        }

        //Método que cambia el formato de la fecha del campo fecha del excel abierto.
        public static String DateConverter(String[,] Array, int indiceFor, int indiceCambio) {
            DateTime dataValue;
            String formatDates = null;

            try {           
                dataValue = DateTime.Parse(Array[indiceFor, indiceCambio]);
                formatDates = dataValue.ToString("yyyy/MM/dd HH:mm:ss");
                
            } catch (Exception e) {
                LOGS.Log("InsercionCOTC (DateConverter) --> " + e.StackTrace + "\n");
            }
            return formatDates;
        }

    }
}
