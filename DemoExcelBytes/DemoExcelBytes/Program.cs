using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoExcelBytes
{
    class Program
    {
        const string fileName = "prueba2.xlsx";
        static void Main(string[] args)
        {
            //Cargar datos (arreglo de bytes)
            var data = GetData();

            //Guardar en archivo Excel (./bin/debug/)
            File.WriteAllBytes(fileName, data);

            //Cargar Objeto Archivo 
            var xlFileInfo = new FileInfo(fileName);

            //Leer Celdas de Excel en memoria
            using (ExcelPackage xlPackage = new ExcelPackage(xlFileInfo))
            {
                //Leer Hoja 1
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[1];
                //Leer los 2 valores de Excel
                Console.WriteLine($"Valor A1: { worksheet.Cells["A1"].Value}");
                Console.WriteLine($"Valor A4: { worksheet.Cells["A4"].Value}");
            }
        }
        public static byte[] GetData()
        {
            return DatosPrueba.Data;
        }
    }
}
