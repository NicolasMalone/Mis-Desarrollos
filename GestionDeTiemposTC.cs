using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Configuration;
using System.Linq;
using System.Windows.Forms;

namespace VolcadoDeAdjuntos {
    class GestionDeTiemposTC {
        static String cred = "SERVER=" + ConfigurationManager.AppSettings.Get("ipbd").ToString() + ";DATABASE=" + ConfigurationManager.AppSettings.Get("database").ToString() + ";UID=" + ConfigurationManager.AppSettings.Get("userbd").ToString() + ";PWD=" + ConfigurationManager.AppSettings.Get("passbd").ToString() + ";";
        static String cred2 = "SERVER=" + ConfigurationManager.AppSettings.Get("ipbd").ToString() + ";DATABASE=" + ConfigurationManager.AppSettings.Get("database").ToString() + ";UID=" + ConfigurationManager.AppSettings.Get("userbd").ToString() + ";PWD=" + ConfigurationManager.AppSettings.Get("passbd").ToString() + ";";
        static String[,] comments;
        static String[,] updateData;
        static int indice = 0;


        public static void ObtenerIndice(DateTime fe) {
            try {
                //Abrimos la conexion a la base de datos utilizando la cadena de conexion.
                MySqlConnection conn = new MySqlConnection(cred);
                MySqlCommand cmd;
                //String dat = DateTime.Now.AddDays(0).ToString("yyyy/MM/dd");

                //Se debe agregar un argumento de tipo DateTime llamado 'fe'
                //fe = fe.AddHours(Double.Parse(DateTime.Now.ToString("HH")));
                //fe = fe.AddMinutes(Double.Parse(DateTime.Now.ToString("mm")));
                //fe = fe.AddSeconds(Double.Parse(DateTime.Now.ToString("ss")));
                String dat = fe.ToString("yyyy/MM/dd");
                MySqlDataReader reader;

                conn.Open();

                cmd = new MySqlCommand("SELECT COUNT(*) From comentario_oceane_tc WHERE DATE_FORMAT(`Lastmodified`, '%Y/%m/%d') = '" + dat + "' AND LOWER(`operationcomment`) like '%[%]%'", conn);
                cmd.CommandTimeout = 2147483;
                reader = cmd.ExecuteReader();

                while (reader.Read()) {
                    comments = new String[Int32.Parse(reader.GetString(0)), 2];
                    indice = Int32.Parse(reader.GetString(0));
                }

                conn.Close();
            } catch (Exception oI) {
                LOGS.Log("GestionDeTiemposTC (ObtenerIndice) --> " + oI + "\n");
            }
        }



        public static void RecuperarDatos(DateTime fe) {
            //Abrimos la conexion a la base de datos utilizando la cadena de conexion.
            MySqlConnection conn = new MySqlConnection(cred);
            MySqlConnection conn2 = new MySqlConnection(cred2);
            MySqlCommand cmd;
            MySqlDataReader reader;
            //String dat = DateTime.Now.AddDays(0).ToString("yyyy/MM/dd");

            //Se debe agregar un argumento de tipo DateTime llamado 'fe'
            //fe = fe.AddHours(Double.Parse(DateTime.Now.ToString("HH")));
            //fe = fe.AddMinutes(Double.Parse(DateTime.Now.ToString("mm")));
            //fe = fe.AddSeconds(Double.Parse(DateTime.Now.ToString("ss")));
            String dat = fe.ToString("yyyy/MM/dd");
            updateData = new string[indice, 3];


            try {
                //Se comprueba si en el campo `comentario` existe el patron '[%]', siendo '%' cualquier caracter.
                cmd = new MySqlCommand("SELECT * From comentario_oceane_tc WHERE DATE_FORMAT(`Lastmodified`, '%Y/%m/%d') = '" + dat + "' AND `operationcomment` like '%[%]%'", conn2);
                conn2.Open();
                cmd.CommandTimeout = 2147483;
                reader = cmd.ExecuteReader();

                long i = 0;
                long total;
                long totalFinal;
                Boolean verfal = false;
                ArrayList listaComenCor = new ArrayList();

                while (reader.Read()) {
                    total = 0;
                    totalFinal = 0;
                    verfal = false;
                    listaComenCor.Clear();
                    long idTicket = reader.GetInt64(0);
                    comments[i, 0] = reader.GetString(2); //Campo ticketID
                    comments[i, 1] = reader.GetString(6); //Campo OperationComment

                    //Si el contenido del campo comentario contiene '`' o  `´` entonces se procede a sustituirlos
                    if (comments[i, 1].Contains("`") || comments[i, 1].Contains("´")) {
                        comments[i, 1] = comments[i, 1].Replace("`", " ");
                        comments[i, 1] = comments[i, 1].Replace("´", " ");
                        comments[i, 1] = comments[i, 1].Replace("<", " ");
                        comments[i, 1] = comments[i, 1].Replace(">", " ");
                    }

                    //try {
                    //Se coge la longitud de la cadena (comentario).
                    for (int j = 0; j < comments[i, 1].Length; j++) {
                        //Se comprueba si en la posicion j de la cadena existe el caracter '['.
                        if (comments[i, 1].ElementAt(j).ToString() == "[") {
                            //Se le pasa el indice a K si en la posicion de j existe un '['.
                            int k = j;
                            //Variable la cual servira para comprobar si hay mas de un digito dentro de '[]'.
                            int secuencia = 0;
                            //Mientras que no haya cierre de corchete el bucle recorrera la cadena a partir del primer valor del corchete de apertura.
                            while (comments[i, 1].ElementAt(k + 1).ToString() != "]") {
                                try {
                                    //Si la secuencia es 0 entonces recogera el valor que haya dentro de los '[]' (cuando hay un solo dígito dentro de los '[]').
                                    if (secuencia == 0) {
                                        total = Int64.Parse(comments[i, 1].Substring(k + 1, 1));
                                        //Si hay mas de un digito dentro de los '[]' entonces se procede a concatenarlos 
                                    } else if (secuencia > 0) {
                                        total = Int64.Parse(total.ToString() + comments[i, 1].Substring(k + 1, 1));
                                    }
                                } catch (Exception) {
                                    verfal = true;
                                    break;
                                }
                                secuencia++;
                                k++;
                                verfal = false;
                            }
                            if (verfal == false) {
                                listaComenCor.Add(total);
                            }
                        }
                    }
                    //} catch (Exception e) {
                    //    LOGS.Log("GestionDeTiemposTC (RecuperarDatos) --> " + e + "\n");
                    //    //MessageBox.Show("" + e);
                    //}

                    foreach (long valor in listaComenCor) {
                        //Se suma todas las cifras del comentario, si hubieran mas de uno.
                        totalFinal = totalFinal + valor;
                    }

                    updateData[i, 0] = totalFinal.ToString();
                    updateData[i, 1] = comments[i, 0];
                    updateData[i, 2] = idTicket.ToString();

                    i++;
                }
                conn2.Close();
            } catch (Exception e) {
                LOGS.Log("GestionDeTiemposTC (RecuperarDatos LastTry) --> " + e + "\n");
                //MessageBox.Show(""+e);
                if (conn2.State == System.Data.ConnectionState.Open) {
                    conn2.Close();
                } else if (conn.State == System.Data.ConnectionState.Open) {
                    conn.Close();
                }
            }

        }


        public static void UpdateTC() {
            MySqlCommand cmdInternal;
            MySqlConnection conn = new MySqlConnection(cred);

            try {
                //MessageBox.Show("Ticket " + comments[i, 0] + " Correcto con el valor de tiempo (mins) --> " + val);
                //commets[i,0] --> Equivale al Ticket de la BD.
                for (int i = 0; i < indice; i++) {
                    cmdInternal = new MySqlCommand("UPDATE comentario_oceane_tc SET `Tiempo` = " + updateData[i, 0] + " WHERE `TicketID` = '" + updateData[i, 1] + "' AND `ID` = " + updateData[i, 2], conn);
                    conn.Open();
                    cmdInternal.CommandTimeout = 2147483;
                    cmdInternal.ExecuteNonQuery();
                    conn.Close();
                }
            } catch (Exception e) {
                LOGS.Log("GestionDeTiemposTC - UpdateTA --> " + e);
                if (conn.State == System.Data.ConnectionState.Open) {
                    conn.Close();
                }
            }
        }



        public static void TiempoFinalTCDA(DateTime fe) {
            MySqlCommand cmdInternal;
            MySqlConnection conn = new MySqlConnection(cred);
            String dat = fe.AddDays(-1).ToString("yyyy/MM/dd");


            try {
                //MessageBox.Show("Ticket " + comments[i, 0] + " Correcto con el valor de tiempo (mins) --> " + val);
                //commets[i,0] --> Equivale al Ticket de la BD.
                cmdInternal = new MySqlCommand("UPDATE tickets_cerrados_da tcda SET tcda.Tiempo = (SELECT SUM(cotc.Tiempo) FROM comentario_oceane_tc cotc WHERE cotc.TicketID = tcda.`Ticket ID` AND DATE_FORMAT(cotc.`Closuredate`, '%Y/%m/%d') = '" + dat + "' GROUP BY cotc.TicketID) WHERE DATE_FORMAT(tcda.`Closure date`, '%Y/%m/%d') = '" + dat + "'", conn);
                conn.Open();
                cmdInternal.CommandTimeout = 2147483;
                cmdInternal.ExecuteNonQuery();
                conn.Close();
                LOGS.Log("Actualizacion --> OK");
            } catch (Exception e) {
                LOGS.Log("GestionDeTiemposTC - UpdateTA --> " + e);
                if (conn.State == System.Data.ConnectionState.Open) {
                    conn.Close();
                }
            }
        }



    }
}
