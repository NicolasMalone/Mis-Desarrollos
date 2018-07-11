using Microsoft.CSharp.RuntimeBinder;
using System;
using System.Configuration;
using System.IO;
using System.Runtime.InteropServices;
using excel = Microsoft.Office.Interop.Excel;

namespace VolcadoDeAdjuntos {
    class LecturaAdjuntos {
        static excel.Application xlApp;
        static excel.Workbooks xlWorkBooks;
        static excel.Workbook xlWorkBook;
        static excel.Worksheet xlWorkSheet;
        static excel.Range range;
        static String[,] dataColRow;
        static String valorCelda;

        public static void LeerFicheroOceane(String nombreAdjunto, DateTime fec) {
            String nFile = ConfigurationManager.AppSettings.Get("ruta").ToString() + "vdata\\" + nombreAdjunto;
            var misValue = Type.Missing;

            //Se comprueba si el archivo adjunto existe en la ruta --> nFile.
            if (!File.Exists(nFile)) {
                return;
            }

            LOGS.Log("Empezando a leer --> "+ nombreAdjunto);
            // abrir el documento 
            xlApp = new excel.Application();
            xlWorkBooks = xlApp.Workbooks;
            xlWorkBook = xlWorkBooks.Open(nFile);

            // seleccion de la hoja de calculo
            xlWorkSheet = (excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // seleccion rango activo
            range = xlWorkSheet.UsedRange;

            // leer las celdas
            int rows = range.Rows.Count;
            int cols = range.Columns.Count;
            dataColRow = new String[rows, cols];

            //Incrementamos valor dependiedo del numero de filas que haya en el excel adjunto.
            int i = 0;
            //Incrementamos valor dependiedo del numero de columnas que haya en el excel adjunto.
            int j = 0;

            //Guardamos el valor de cada celda en uso del adjunto en un array bidi.
            for (int row = 5; row <= rows; row++) {
                j = 0;
                for (int col = 1; col <= cols; col++) {
                    try {
                        // Valor de la celda actual.
                        valorCelda = (range.Cells[row, col] as excel.Range).Value.ToString();
                        dataColRow[i, j] = valorCelda;
                        j++;
                        //MessageBox.Show(valorCelda);
                    } catch (RuntimeBinderException) {
                        dataColRow[i, j] = "";
                        j++;
                    }
                }
                i++;
            }

            //Seleccionaremos el método según el nombre del adjunto.
            if (nombreAdjunto.ToLower().Contains("COMENTARIOS OCEANE- TICKETS CERRADOS - POSTVENTA CONECTA PYMES".ToLower())) {
                LOGS.Log("Lectura del adjunto " + nombreAdjunto+ " --> OK");
                InsercionCOTC.InsercionData(dataColRow, rows, fec);
            } else if (nombreAdjunto.ToLower().Contains("CERTIFICACION POSTVENTA REVISION DIARIA".ToLower())) {
                LOGS.Log("Lectura del adjunto " + nombreAdjunto + " --> OK");
                InsercionTCDA.InsercionData(dataColRow, rows, fec);
            } else if (nombreAdjunto.ToLower().Contains("BACKLOG OCEANE (ibermatica)".ToLower())) {
                LOGS.Log("Lectura del adjunto " + nombreAdjunto + " --> OK");
                InsercionBO.InsercionData(dataColRow, rows, fec);
            } else if (nombreAdjunto.ToLower().Contains("COMENTARIOS OCEANE- TICKETS ABIERTOS - POSTVENTA CONECTA PYMES".ToLower())) {
                LOGS.Log("Lectura del adjunto " + nombreAdjunto + " --> OK");
                InsercionCOTA.InsercionData(dataColRow, rows, fec);
            }

            //Cerramos el worbook. 
            xlWorkBook.Close(false, misValue, misValue);
            //Cerramos el conjunto de workbooks
            xlWorkBooks.Close();
            //Salimos de excel.
            xlApp.Quit();
            //Liberamos los tres objetos usados anteriomente.
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlWorkBooks);
            Marshal.ReleaseComObject(xlApp);
        }

    }
}
