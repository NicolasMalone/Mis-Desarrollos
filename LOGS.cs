using System;
using System.IO;

namespace VolcadoDeAdjuntos {
    class LOGS {

        public static void Log(String mensaje) {
            //Boolean error;
            String path;
            int contador = 0;
            //do {
                //error = false;
                path = AppDomain.CurrentDomain.BaseDirectory + "\\" + "\\logs\\log" + DateTime.Now.ToString("ddMMyyyy") + "-" + contador.ToString() + ".txt";
                try {
                    File.AppendAllLines(path, new String[] { DateTime.Now.ToString("dd/MM/yyyy -- HH:mm:ss  ||  ") + mensaje });
                    Console.WriteLine(mensaje);
                } catch (IOException e) {
                    Console.WriteLine(e.Message.ToString());
                    //error = true;
                    contador++;
                }
            //} while (error);
        }


    }
}
