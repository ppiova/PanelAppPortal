using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LecturaDatosExcel_ExcelDataReader
{
    public class ExcelReader
    {

        public List<DataTable> readExcel(string filePath)
        {

            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader;

            //1. Reading Excel file
            if (Path.GetExtension(filePath).ToUpper() == ".XLS")
            {
                //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
            }
            else
            {
                //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }

            //2. DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet result = excelReader.AsDataSet();


            //Almacena los datatables de cada una de las tablas del archivo
            List<DataTable> DataTablesList = new List<DataTable>();

            //Cantidad de tablas (Se crea una por cada hoja del archivo)
            var tablesQuantities = result.Tables.Count;

            //Creamos un DataTable por cada tabla
            for (int i = 0; i < tablesQuantities; i++)
            {
                DataTablesList.Add(result.Tables[i]);
            }

            excelReader.Close();

            return DataTablesList;
        }

    }
}


//EJEMPLO DE LLAMADO AL MÉTODO DE LECTURA Y MAPEO DE HOJAS DE DATOS DEL ARCHIVO 
//Ejecutar la lectura del archivo (podemos leer formato .xls y .xlsx)
//string filePath = @"C:\Users\Federico Kloster\source\repos\LecturaDatosExcel-ExcelDataReader\LecturaDatosExcel-ExcelDataReader\PruebaLectura.xlsx";
//var reader = new ExcelReader();
//var DataTablesList = reader.readExcel(filePath);

//var EntidadHojaUnoList = new List<EntidadHojaUno>();
//var EntidadHojaDosList = new List<EntidadHojaDos>();

//            //Recorremos las tablas del archivo, mapeando cada una con su modelo correspondiente.
//            for (int i = 0; i<DataTablesList.Count; i++)
//            {
       
//                var dataTable = DataTablesList[i];

//                switch (i)
//                {
//                    case 0:
//                        for (int j = 1; j<dataTable.Rows.Count; j++)
//                        {
//                             DataRow row = dataTable.Rows[j];

//EntidadHojaUnoList.Add(new EntidadHojaUno()
//{
//    Id = row[0].ToString(),
//                                Nombre = row[1].ToString(),
//                                Apellidos = row[2].ToString()
//                            });
//                        }
//                        break;
//                    case 1:
//                        for (int j = 1; j<dataTable.Rows.Count; j++)
//                        {
//                            DataRow row = dataTable.Rows[j];

//EntidadHojaDosList.Add(new EntidadHojaDos()
//{
//    Id = row[0].ToString(),
//                                Direccion = row[1].ToString(),
//                                Telefono = row[2].ToString(),
//                                Edad = row[3].ToString()
//                            });
//                        }
//                        break;
//                    default:
//                        break;
//                }
//            }



