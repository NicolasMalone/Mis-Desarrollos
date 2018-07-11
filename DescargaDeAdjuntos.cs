using System;
using System.Collections;
using System.Configuration;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace VolcadoDeAdjuntos {
    class DescargaDeAdjuntos {
        static Outlook.Application ok = new Outlook.Application();
        static Outlook.NameSpace ns = ok.GetNamespace("MAPI");
        static Outlook.MAPIFolder mailboxFolder;
        static Outlook.MAPIFolder folder;
        static Outlook.Stores stores;
        static Outlook.Store store;
        static Outlook.Folders folders;
        //string folderList;
        static int comp = 0;
        static ArrayList adjuntos = new ArrayList();
        static String[] extensionsArray = { ".pdf", ".doc", ".docx", ".xlsx", ".xls", ".ppt", ".cs", ".vsd", ".zip", ".rar", ".txt", ".csv", ".proj", ".jpg", ".jpeg", ".png" };

        private static void Verificar_Correos(string directory_name, DateTime fec) {
            bool verfal;
            verfal = Existe_Buzon();
            if ((verfal == true)) {
                ns = ok.Session;
                stores = ns.Stores;
                // cnt debe empezar desde 1 (indice 0 no existe).
                int cnt = 1;
                foreach (Outlook.Store store in ns.Stores) {
                    // Si se quiere buscar otro buzon, se debe cambiar el valor en App.Config
                    if (store.DisplayName.ToLower().Contains(ConfigurationManager.AppSettings.Get("busqueda_buzon").ToLower().ToString())) {
                        //MessageBox.Show(store.DisplayName.ToString());
                        break;
                    } else {
                        // Hasta que no entre en el if cnt se incremente (cnt representa las carpetas del buzon por indices).
                        cnt++;
                        // MsgBox(store.DisplayName.ToString)
                    }
                }

                // stores(cnt) es el buz�n el cual se buscar�n los correos a descargar.
                store = stores[cnt];
                mailboxFolder = store.GetRootFolder();
                folders = mailboxFolder.Folders;
                for (int j = 1; (j <= folders.Count); j++) {
                    folder = folders[j];
                    // folderList += folder.Name + Environment.NewLine
                    if ((folder.Name.ToLower().Equals(directory_name.ToLower()))) {
                        // MsgBox(folder.Name, , "Carpeta de Busqueda")
                        SeleccionarCarpeta(folder, fec);
                        j = folders.Count;
                    }

                }

                // MessageBox.Show(folderList, "Lista Outlook")
                //folderList = null;
                Marshal.ReleaseComObject(folder);
            }

        }


        public static bool Existe_Buzon() {
            string nombre_buzon_cliente;
            string nombre_buzon_largo;
            bool Existe_Buzon = false;
            // Guardamos el nombre del buzon.
            nombre_buzon_cliente = ConfigurationManager.AppSettings.Get("buzon").ToString();
            nombre_buzon_largo = "Buzón - " + nombre_buzon_cliente;
            //MessageBox.Show(nombre_buzon_largo);

            foreach (Outlook.MAPIFolder mailboxFolder in ns.Folders) {
                if (mailboxFolder.Name.ToLower().Contains(nombre_buzon_largo.ToLower())) {
                    Existe_Buzon = true;
                }

            }
            return Existe_Buzon;
        }


        private static void SeleccionarCarpeta(Outlook.MAPIFolder folder, DateTime fec) {
            Outlook.Folders childFolders = folder.Folders;
            if ((childFolders.Count > 0)) {
                // A partir del root se hace una busqueda para ver que carpetas hay.
                foreach (Outlook.Folder childFolder in childFolders) {
                    // Se busca la carpeta deseada, para poder mostrar sus mensajes mas adelante.
                    if (childFolder.FolderPath.Contains(ConfigurationManager.AppSettings.Get("carpeta_a_buscar").ToLower().ToString())) {
                        // Console.WriteLine(childFolder.FolderPath)
                        // Se buscan mas carpetas dentro de esta carpeta.
                        SeleccionarCarpeta(childFolder, fec);
                    }

                }

            }

            // Mensaje para comprobar la busqueda.   
            //MessageBox.Show("Buscando por items en.. " + folder.FolderPath);
            ArchivosAdjuntos(folder, (ConfigurationManager.AppSettings.Get("ruta").ToString() + "vdata\\"), fec);
        }


        private static void ArchivosAdjuntos(Outlook.MAPIFolder folder, string ruta, DateTime fec) {
            Outlook.Items fi = folder.Items;
            // Se ordenan los correos por fecha de entrega para evitar leer todos los correos que no sean del d�a actual.
            fi.Sort("[ReceivedTime]", true);
            if ((fi.ToString() != null)) {
                foreach (Outlook.MailItem item in fi) {
                    Outlook.MailItem mi = item;
                    // Se listan todos los adjuntos del buzon.
                    Outlook.Attachments attachments = mi.Attachments;
                    string nombreRemitente;
                    DateTime fecha;
                    // Se aplica un margen minimo y maximo el cual busque los correos.
                    DateTime intervaloMenor = fec.AddHours(1);
                    DateTime intervaloMayor = fec.AddHours(12);
                    fecha = mi.ReceivedTime;
                    nombreRemitente = mi.SenderName;
                    //fechaHora = fecha.Split(' ');
                    //MessageBox.Show(fecha.Date.ToString());
                    if (nombreRemitente.ToLower().Contains(ConfigurationManager.AppSettings.Get("remitente").ToLower()) && fecha.Date.Equals(fec.Date) && (mi.Subject.ToLower().Contains("CERTIFICACION POSTVENTA DIARIA".ToLower()) || mi.Subject.ToLower().Contains("BACKLOG OCEANE (ibermatica)".ToLower()) || mi.Subject.ToLower().Contains("COMENTARIOS OCEANE- TICKETS CERRADOS - POSTVENTA CONECTA PYMES".ToLower()) || mi.Subject.ToLower().Contains("COMENTARIOS OCEANE- TICKETS ABIERTOS - POSTVENTA CONECTA PYMES".ToLower())) && fecha.Hour >= intervaloMenor.Hour && fecha.Hour <= intervaloMayor.Hour) {
                        // Comprobamos si hay algun adjunto dentro del email recibido.
                        if ((attachments.Count != 0)) {
                            for (int i = 1; (i <= mi.Attachments.Count); i++) {
                                // Guardamos el nombre del adjunto en "fn"
                                string fn = mi.Attachments[i].FileName.ToLower();
                                for (int j = 0; (j <= (extensionsArray.Length - 1)); j++) {
                                    // Utilizamos el array de extensiones para comprobar su compatibilidad y poder descargarlo.
                                    if (fn.Contains(extensionsArray[j].ToString())) {
                                        //  MessageBox.Show("Dia: " + fechaHora(0) + " Hora: " + fechaHora(1) & vbNewLine & "Adjunto: " + mi.Attachments(i).FileName, "Dato Adjunto")                                         
                                        //Se realiza la comprobacion del archivo descargado (El primer if puede ser confuso, pero se realiza para evitar hacer una comrpobacion o insercion errónea de datos).
                                        if (mi.Attachments[i].FileName.Contains("Tickets Oceane Cerrados dia anterior - P.Segundo nivel")) {
                                        } else {
                                            if (mi.Attachments[i].FileName.ToLower().Contains("COMENTARIOS OCEANE- TICKETS CERRADOS - POSTVENTA CONECTA PYMES".ToLower())) {
                                                LOGS.Log("InsercionCOTA (InsercionData) --> Se Descarga el fichero 'COMENTARIOS OCEANE- TICKETS CERRADOS - POSTVENTA CONECTA PYMES'");
                                                // Se guarda el documento en la ruta especificada (se especifica en App.Config). 
                                                mi.Attachments[i].SaveAsFile(ruta + "" + mi.Attachments[i].FileName);
                                                LecturaAdjuntos.LeerFicheroOceane(mi.Attachments[i].FileName, fec);
                                            } else if (mi.Attachments[i].FileName.ToLower().Contains("CERTIFICACION POSTVENTA REVISION DIARIA".ToLower())) {
                                                LOGS.Log("InsercionCOTA (InsercionData) --> Se Descarga el fichero 'Certificacion Postventa'");
                                                //Se guarda el documento en la ruta especificada (se especifica en App.Config). 
                                                mi.Attachments[i].SaveAsFile(ruta + "" + mi.Attachments[i].FileName);
                                                LecturaAdjuntos.LeerFicheroOceane(mi.Attachments[i].FileName, fec);
                                            } else if (mi.Attachments[i].FileName.ToLower().Contains("BACKLOG OCEANE (ibermatica)".ToLower())) {
                                                if (comp < 1) {
                                                    LOGS.Log("InsercionCOTA (InsercionData) --> Se Descarga el fichero 'BACKLOG OCEANE (ibermatica)'");
                                                    // Se guarda el documento en la ruta especificada (se especifica en App.Config). 
                                                    mi.Attachments[i].SaveAsFile(ruta + "" + mi.Attachments[i].FileName);
                                                    LecturaAdjuntos.LeerFicheroOceane(mi.Attachments[i].FileName, fec);
                                                    comp++;
                                                }
                                            } else if (mi.Attachments[i].FileName.ToLower().Contains("COMENTARIOS OCEANE- TICKETS ABIERTOS - POSTVENTA CONECTA PYMES".ToLower())) {
                                                LOGS.Log("InsercionCOTA (InsercionData) --> Se Descarga el fichero 'COMENTARIOS OCEANE- TICKETS ABIERTOS - POSTVENTA CONECTA PYMES'");
                                                // Se guarda el documento en la ruta especificada (se especifica en App.Config). 
                                                mi.Attachments[i].SaveAsFile(ruta + "" + mi.Attachments[i].FileName);
                                                LecturaAdjuntos.LeerFicheroOceane(mi.Attachments[i].FileName, fec);
                                            }
                                            j = extensionsArray.Length;
                                        }

                                    }

                                }

                            }

                        }
                    } else if (fecha.Date < fec.Date) {
                        break;
                    }
                }


            }

        }

        public static void PasoFecha(DateTime fec) {
            // Si se desea cambiar la carpeta donde buscar los adjuntos, se debera cambiar el valor en App.Config             
            Verificar_Correos(ConfigurationManager.AppSettings.Get("carpeta_a_buscar").ToString(), fec);
        }
    }
}
