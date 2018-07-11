using System;
using System.Configuration;
using System.IO;
using System.Threading;

namespace VolcadoDeAdjuntos {
    class Program {

        static String path = (ConfigurationManager.AppSettings.Get("ruta").ToString() + "vdata\\log.txt").ToString();

        static void Main(string[] args) {
            //Bucle para incertar por dias (custom).
            //for (int i = 31; i <= 31; i++) {
            //DateTime a = DateTime.Parse(i + "/05/2018");

            try {
                // Si la carpeta tempcm existe se borra y se vuelve a crear.
                if (Directory.Exists((ConfigurationManager.AppSettings.Get("ruta").ToString() + "vdata\\"))) {
                    Directory.Delete((ConfigurationManager.AppSettings.Get("ruta").ToString() + "vdata\\"), true);
                    Directory.CreateDirectory((ConfigurationManager.AppSettings.Get("ruta").ToString() + "vdata\\"));

                    //Desde la clase deseada llamar a este metodo pasandole como parametro una fecha (DateTime).
                    if (ConfigurationManager.AppSettings.Get("updateOnly").ToString() == "1") {
                        DescargaDeAdjuntos.PasoFecha(DateTime.Today);
                        //DescargaDeAdjuntos.PasoFecha(a);
                        Thread.Sleep(2000);
                    }
                    LOGS.Log("Actualizando campo 'Tiempo' de COTC...");
                    GestionDeTiemposTC.ObtenerIndice(DateTime.Today);
                    GestionDeTiemposTC.RecuperarDatos(DateTime.Today);
                    GestionDeTiemposTC.UpdateTC();
                    GestionDeTiemposTC.TiempoFinalTCDA(DateTime.Today);
                    Thread.Sleep(2000);
                    LOGS.Log("Actualizando campo 'Tiempo' de COTA...");
                    GestionDeTiemposTA.ObtenerIndice(DateTime.Now);
                    GestionDeTiemposTA.RecuperarDatos(DateTime.Now);
                    GestionDeTiemposTA.UpdateTA();
                } else {
                    // Si la carpeta tempcm no existe se crea.
                    Directory.CreateDirectory((ConfigurationManager.AppSettings.Get("ruta").ToString() + "vdata\\"));

                    //Desde la clase deseada llamar a este metodo pasandole como parametro una fecha (DateTime).
                    if (ConfigurationManager.AppSettings.Get("updateOnly").ToString() == "1") {
                        DescargaDeAdjuntos.PasoFecha(DateTime.Today);
                        //DescargaDeAdjuntos.PasoFecha(a);
                        Thread.Sleep(2000);
                    }

                    LOGS.Log("Actualizando campo 'Tiempo' de COTC...");
                    GestionDeTiemposTC.ObtenerIndice(DateTime.Today);
                    GestionDeTiemposTC.RecuperarDatos(DateTime.Today);
                    GestionDeTiemposTC.UpdateTC();
                    GestionDeTiemposTC.TiempoFinalTCDA(DateTime.Today);
                    Thread.Sleep(2000);
                    LOGS.Log("Actualizando campo 'Tiempo' de COTA...");
                    GestionDeTiemposTA.ObtenerIndice(DateTime.Today);
                    GestionDeTiemposTA.RecuperarDatos(DateTime.Today);
                    GestionDeTiemposTA.UpdateTA();
                }
                LOGS.Log("Ejecucion Terminada --> OK");
            } catch (IOException e) {
                LOGS.Log("Catch General 1 --> " + e.StackTrace);
                //MessageBox.Show(e.StackTrace, "Advertencia", MessageBoxButtons.OK);
            } catch (Exception ex) {
                LOGS.Log("Catch General 2 --> " + ex.StackTrace);
                //MessageBox.Show(ex.StackTrace, "Advertencia", MessageBoxButtons.OK);
            }
        }
        //}


    }
}