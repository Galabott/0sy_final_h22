using ClosedXML.Excel;
using ExcelToExcel.ViewModels;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelToExcel.Models
{
    public class EspeceXL
    {
        private string sheetname = "especes";

        private IXLWorkbook wb;

        public string Filename { get; set; }
        public bool IsReadOnly { get; private set; } = false;

        public string CSVSeparator { get; set; } = ";";


        public EspeceXL(string filename)
        {
            Filename = filename;
        }

        public void LoadFile()
        {
            wb = new XLWorkbook(Filename);
            IsReadOnly = false;
        }

        public void LoadFileReadOnly()
        {
            using (var fs = new FileStream(Filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                wb = new XLWorkbook(fs);
                IsReadOnly = true;
            }
        }

        public object GetCell(int sheet, int row, int column)
        {
            return wb.Worksheet(sheet).Cell(row, column).Value;
        }

        public object GetCell(string sheet, int row, int column)
        {
            return wb.Worksheet(sheet).Cell(row, column).Value;
        }

        /// <summary>
        /// Retourne le contenu en format CSV
        /// </summary>
        /// <returns>Format CSV</returns>
        public string GetCSV()
        {
            if (!validFileContent())
            {
                throw new ArgumentException("Mauvais format de fichier!");
            }

            string result = string.Empty;

            var worksheet = wb.Worksheet(sheetname);
            var lastCellAddress = worksheet.RangeUsed().LastCell().Address;

            var res = worksheet.Rows(1, lastCellAddress.RowNumber)
                .Select(r => string.Join(CSVSeparator, r.Cells(1, lastCellAddress.ColumnNumber)
                        .Select(cell =>
                        {
                            var cellValue = cell.GetValue<string>();
                            return cellValue.Contains(CSVSeparator) ? $"\"{cellValue}\"" : cellValue;
                        })));

            foreach (var r in res)
            {
                result += r + Environment.NewLine;
            }

            return result;
        }

        public List<Espece> GetAsList()
        {
            List<Espece> result = new List<Espece>();
            int startRow = 2;

            var worksheet = wb.Worksheet(sheetname);
            var lastCellAddress = worksheet.RangeUsed().LastCell().Address;

            result = worksheet.Rows(startRow, lastCellAddress.RowNumber)
                .Select(r =>
                    new Espece { 
                        Nom = (string)(r.Cell(1).Value), 
                        NomLatin = (string)(r.Cell(2).Value),
                        Habitat = (string)(r.Cell(3).Value),
                    }
                ).ToList();

            return result;
        }

        /// <summary>
        /// Permet de valider le contenu du fichier "Espece"
        /// </summary>
        /// <returns></returns>
        private bool validFileContent()
        {
            
            var valid = wb.Worksheets.Contains(sheetname);

            if (valid)
            {
                var ws = wb.Worksheet(sheetname);
                var nom = ws.Cell(1, 1).Value.ToString().Trim().ToLower();
                var nomLatin = ws.Cell(1, 2).Value.ToString().Trim().ToLower();
                var habitat = ws.Cell(1, 3).Value.ToString().Trim().ToLower();

                valid &= nom == "nom";
                valid &= nomLatin == "nom latin";
                valid &= habitat == "habitat";
            }

            return valid;
        }

        /// <summary>
        /// Permet de sauvegarder le contenu dans un fichier CSV
        /// </summary>
        /// <param name="filename"></param>
        public void SaveCSV(string filename)
        {
            /// xTODO : Q05 : Ajouter les validations pour passer les tests
            /// 
            if (filename[^3..] == "csv" && !VerifySpecialCharsInString(filename))
            {
                var output = GetCSV();

                using (var writer = new StreamWriter(filename, false, System.Text.Encoding.UTF8))
                {
                    writer.Write(output);
                }
            }
            else
            {
                throw new ArgumentException("Mauvais format de fichier!");
            }
        }

        public void SaveJson(string filename)
        {
            /// xTODO : Q06 Ajouter les validations pour passer les tests
            /// 
            if (filename[^4..] == "json" && !VerifySpecialCharsInString(filename))
            {
                var lst = GetAsList();

                string output = JsonConvert.SerializeObject(lst, Formatting.Indented);

                using (var writer = new StreamWriter(filename))
                {
                    writer.Write(output);
                }
            }
            else
            {
                throw new ArgumentException("Mauvais format de fichier!");

            }

        }

        public void SaveToFile(string filename, bool overwrite = false)
        {
            /// xTODO : Q08 Ajouter les validations pour passer les tests
            
            var ext = Path.GetExtension(filename).ToLower();

            switch (ext)
            {
                case ".csv":
                    SaveCSV(filename);
                    break;
                case ".json":
                    SaveJson(filename);
                    break;
                case ".xlsx":
                    SaveXls(filename);
                    break;
                default:
                    /// xTODO : Q09 Lancer l'exception ArgumentException avec le message "Type inconnu" et le nom du paramètre filename
                    /// 
                    throw new ArgumentException("Format inconnu!");
                    break;
            }
        }

        public void SaveXls(string filename)
        {
            /// xTODO : Q07 Ajouter les validations pour passer les tests
            if (filename[^4..] == "xlsx" && !VerifySpecialCharsInString(filename))
            {
                wb.SaveAs(filename);
            }
            else
            {
                throw new ArgumentException("Mauvais format de fichier!");
            }

        }

        private bool VerifySpecialCharsInString(string s)
        {
            Regex r = new Regex("[a-z][A-Z][.$/][0-9]");
            if (r.IsMatch(s))
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }
}
