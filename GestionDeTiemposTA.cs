using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Configuration;
using System.Linq;
using System.Windows.Forms;

namespace VolcadoDeAdjuntos {
    class GestionDeTiemposTA {

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
                cmd = new MySqlCommand("SELECT COUNT(*) From comentarios_oceane_ta WHERE DATE_FORMAT(`Lastmodified`, '%Y/%m/%d') = '" + dat + "' AND `operationcomment` like '%[%]%'", conn);
                cmd.CommandTimeout = 2147483;
                reader = cmd.ExecuteReader();

                while (reader.Read()) {
                    indice = Int32.Parse(reader.GetString(0));
                    comments = new String[Int32.Parse(reader.GetString(0)), 2];
                }

                conn.Close();
            } catch (Exception e) {
                LOGS.Log("GestionDeTiemposTA (RecogerIndice) --> " + e + "\n");
            }
        }



        public static void RecuperarDatos(DateTime fe) {
            //Abrimos la conexion a la base de datos utilizando la cadena de conexion.
            MySqlConnection conn2 = new MySqlConnection(cred2);
            MySqlCommand cmd;
            MySqlDataReader reader;
            long i = 0;
            long total = 0;
            long totalFinal = 0;
            Boolean verfal = false;
            ArrayList listaComenCor = new ArrayList();
            //String dat = DateTime.Now.AddDays(0).ToString("yyyy/MM/dd");


            //Se debe agregar un argumento de tipo DateTime llamado 'fe'
            //fe = fe.AddHours(Double.Parse(DateTime.Now.ToString("HH")));
            //fe = fe.AddMinutes(Double.Parse(DateTime.Now.ToString("mm")));
            //fe = fe.AddSeconds(Double.Parse(DateTime.Now.ToString("ss")));
            String dat = fe.ToString("yyyy/MM/dd");
            updateData = new string[indice, 3];



            try {
                //Se comprueba si en el campo `comentario` existe el patron '[%]', siendo '%' cualquier caracter.
                cmd = new MySqlCommand("SELECT ID, TicketID, OperationComment From comentarios_oceane_ta WHERE DATE_FORMAT(`Lastmodified`, '%Y/%m/%d') = '" + dat + "' AND OperationComment like '%[%]%' ORDER BY `TicketID`, CommentNumber desc", conn2);
                conn2.Open();
                cmd.CommandTimeout = 2147483;
                reader = cmd.ExecuteReader();

                while (reader.Read()) {

                    total = 0;
                    totalFinal = 0;
                    verfal = false;
                    listaComenCor.Clear();                  
                    long idTicket = reader.GetInt64(0);
                    comments[i, 0] = reader.GetString(1); //Campo ticketID
                    comments[i, 1] = reader.GetString(2); //Campo OperationComment

                    //Si el contenido del campo comentario contiene '`' o  '´' o '\' entonces se procede a sustituirlos
                    if (comments[i, 1].Contains("`") || comments[i, 1].Contains("´") || comments[i, 1].Contains("\"") || comments[i, 1].Contains("_")) {
                        comments[i, 1] = comments[i, 1].Replace("`", " ");
                        comments[i, 1] = comments[i, 1].Replace("´", " ");
                        comments[i, 1] = comments[i, 1].Replace("\"", " ");
                        //comments[i, 1] = comments[i, 1].Replace(".", " ");
                        comments[i, 1] = comments[i, 1].Replace("_", " ");
                    }

                    if (comments[i, 1].Contains("<br/>")) {
                        comments[i, 1] = comments[i, 1].Replace("<br/>", " ");
                    }

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
                                    secuencia++;
                                    k++;
                                    verfal = false;
                                } catch (FormatException) {
                                    verfal = true;
                                    break;
                                }
                            }
                            if (verfal == false) {
                                listaComenCor.Add(total);
                            }
                        }
                    }

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
                LOGS.Log("GestionDeTiemposTA (LastTry) --> " + e + "\n");
                //MessageBox.Show("GestionDeTiemposTA (LastTry)" + e);
                if (conn2.State == System.Data.ConnectionState.Open) {
                    conn2.Close();
                }
            }
            //}

        }

        public static void UpdateTA() {
            MySqlCommand cmdInternal;
            MySqlConnection conn = new MySqlConnection(cred);

            try {
                //MessageBox.Show("Ticket " + comments[i, 0] + " Correcto con el valor de tiempo (mins) --> " + val);
                //commets[i,0] --> Equivale al Ticket de la BD.
                for (int i = 0; i < indice; i++) {
                    cmdInternal = new MySqlCommand("UPDATE comentarios_oceane_ta SET Tiempo = " + Int64.Parse(updateData[i, 0]) + " WHERE `TicketID` = '" + updateData[i, 1] + "' AND ID = " + Int64.Parse(updateData[i, 2]), conn);
                    conn.Open();
                    cmdInternal.CommandTimeout = 2147483;
                    cmdInternal.ExecuteNonQuery();
                    conn.Close();
                }
                LOGS.Log("Actualiacion de Campos --> OK");
            } catch (Exception e) {
                LOGS.Log("GestionDeTiemposTA - UpdateTA --> " + e);
                if (conn.State == System.Data.ConnectionState.Open) {
                    conn.Close();
                }
            }
        }


    }
}
