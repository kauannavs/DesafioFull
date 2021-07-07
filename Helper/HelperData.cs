using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DesafioRPA.Helper
{
    public class HelperData
    {
        #region Define caminho dos diretorios utilizados
        public static string directoryPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\\..\\..\\"));
        public static string fileNameSource = $"{directoryPath}Lista_de_CEPs - DESAFIO RPA.xlsx";
        public static string fileNameTarget = $"{directoryPath}Resultado.xlsx";
        public static string fileNameLog = $"{directoryPath}Log\\Log.txt";
        public static string fileNameCepRequest = $"{directoryPath}Log\\LogRequisicoesPorCep.txt";
        #endregion

        public static void Log(string fileName, string logMessage)
        {
            using (StreamWriter w = File.AppendText(fileName))
            {
                w.WriteLine($" {DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")} - {logMessage}");
                Console.WriteLine(logMessage);
            }
        }
        public static void WriteExceptionLog(string name, string message)
        {
            Log(fileNameLog, $"----------------------------------");
            Log(fileNameLog, $"Error - {name} - ");
            Log(fileNameLog, $"Error - {message} - ");
            Log(fileNameLog, $"----------------------------------");
        }
        public static List<(int, int)> ExtractData()
        {
            Log(fileNameLog, $"ExtractData - Start");
            List<(int, int)> listExtractedData = new List<(int, int)>();
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage package = new ExcelPackage(new FileInfo(fileNameSource)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    int countColumn = worksheet.Dimension.End.Column;
                    int countRow = worksheet.Dimension.End.Row;

                    for (int row = 2; row <= countRow; row++)
                    {
                        listExtractedData.Add((
                            Convert.ToInt32(worksheet.Cells[row, 2].Value.ToString()),
                            Convert.ToInt32(worksheet.Cells[row, 3].Value.ToString())));
                    }
                }

                Log(fileNameLog, $"ExtractData - Stop");
                return RemoveDuplicateCep(listExtractedData);
            }
            catch (InvalidOperationException ex)
            {
                Log(fileNameLog, $"----------------------------------");
                Log(fileNameLog, $"Error - {ex.GetType().FullName} - ");
                Log(fileNameLog, $"Error - {ex.Message} - ");
                Log(fileNameLog, $"----------------------------------");
            }
            return null;
        }
        public static List<(int, int)> RemoveDuplicateCep(List<(int, int)> listData)
        {
            if (listData == null)
                return null;

            try
            {
                listData.Sort();
                List<(int, int)> listCepRange = new List<(int, int)>();
                for (int row = 0; row <= listData.Count - 1; row++)
                {
                    if (row == 0)
                        listCepRange.Add((listData[row].Item1, listData[row].Item2));
                    else
                    {
                        var prevCepInicial = listData[row - 1].Item1;
                        var prevCepFinal = listData[row - 1].Item2;
                        if (!(prevCepInicial == listData[row].Item1) || !(prevCepFinal == listData[row].Item2))
                            listCepRange.Add((listData[row].Item1, listData[row].Item2));
                    }
                }
                return listCepRange;
            }
            catch (InvalidOperationException ex)
            {
                WriteExceptionLog(ex.GetType().FullName, ex.Message);
            }
            return null;
        }
        public static void CreateResultFile()
        {
            try
            {
                if (!File.Exists(fileNameTarget))
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (var ExcelPkg = new ExcelPackage(new FileInfo(fileNameTarget)))
                    {
                        var sheet = ExcelPkg.Workbook.Worksheets.Add("Resultado");
                        //Fill title row
                        sheet.Cells[1, 1].Value = "Logradouro/Nome";
                        sheet.Cells[1, 2].Value = "Bairro/Distrito";
                        sheet.Cells[1, 3].Value = "Localidade/UF";
                        sheet.Cells[1, 4].Value = "CEP";
                        sheet.Cells[1, 5].Value = "Data/hora/minuto";
                        sheet.Cells["A1:E1"].Style.Font.Bold = true;

                        ExcelPkg.Save();
                    }
                }
            }
            catch (InvalidOperationException ex)
            {
                WriteExceptionLog(ex.GetType().FullName, ex.Message);
            }
        }
        public static void WriteResultFile(InfoLocation infoLocation)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var ExcelPkg = new ExcelPackage(new FileInfo(fileNameTarget)))
                {
                    ExcelWorksheet sheet = ExcelPkg.Workbook.Worksheets.FirstOrDefault();
                    var newRow = sheet.Dimension.End.Row + 1;

                    sheet.Cells[newRow, 1].Value = infoLocation.Logradouro;
                    sheet.Cells[newRow, 2].Value = infoLocation.Bairro;
                    sheet.Cells[newRow, 3].Value = infoLocation.LocalidadeUF;
                    sheet.Cells[newRow, 4].Value = infoLocation.CEP;
                    sheet.Cells[newRow, 5].Value = infoLocation.Data;

                    ExcelPkg.Save();
                }
            }
            catch (InvalidOperationException ex)
            {
                WriteExceptionLog(ex.GetType().FullName, ex.Message);
            }
        }
    }
}
