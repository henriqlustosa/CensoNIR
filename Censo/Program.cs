using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;


namespace Censo
{
    class Program
    {

        public static DateTime SafeConvertToDateTime(object value, DateTime defaultValue)
        {
            return value != DBNull.Value ? Convert.ToDateTime(value): defaultValue;
        }
        private const string URL = "http://intranethspm:5003/hspmsgh-api/censoNepi/";

        static void Main()
        {
            DateTime today = DateTime.Now;
            string excelFilePath = $"\\\\hspmins2\\CensoHenriqueSGH\\Censo_{today:yyyy_MM_dd_HH_mm_ss}.xlsx";

            try
            {
                List<Censo> censos = ObterDadosCenso();
                if (censos == null || censos.Count == 0)
                {
                    Console.WriteLine("Nenhum dado encontrado.");
                    return;
                }


                ExportDataTableToExcel(CriarDataTable(censos), excelFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro: {ex.Message}");
            }
        }

        private static List<Censo> ObterDadosCenso()
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);
                request.Method = "GET";

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    string responseText = reader.ReadToEnd();
                    List<Censo> censos = JsonConvert.DeserializeObject<List<Censo>>(responseText);

                    foreach (var censo in censos)
                    {
                        censo.tempo = TratarTempo(censo.tempo);
                    }

                    return censos;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro ao buscar os dados: {ex.Message}");
                return new List<Censo>();
            }
        }

        private static System.Data.DataTable CriarDataTable(List<Censo> censos)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            if (censos.Count == 0) return dt;

            dt.Columns.Add("nr_quarto");
            dt.Columns.Add("nm_unidade_funcional");
            dt.Columns.Add("dt_internacao_data");
            dt.Columns.Add("dt_internacao_hora");
            dt.Columns.Add("cd_prontuario");
            dt.Columns.Add("nm_paciente");
            dt.Columns.Add("in_sexo");
            dt.Columns.Add("nr_idade");
            dt.Columns.Add("dt_nascimento");
            dt.Columns.Add("vinculo");
            dt.Columns.Add("nm_especialidade");
            dt.Columns.Add("nm_medico");
            dt.Columns.Add("cod_CID");
            dt.Columns.Add("descricaoCID");
            dt.Columns.Add("tempo");
            dt.Columns.Add("nm_origem");
            dt.Columns.Add("dt_ultimo_evento_data");
            dt.Columns.Add("dt_ultimo_evento_hora");
            dt.Columns.Add("nr_convenio");

            foreach (var censo in censos)
            {
                dt.Rows.Add(
                    censo.nr_quarto, censo.nm_unidade_funcional, censo.dt_internacao_data,
                    censo.dt_internacao_hora, censo.cd_prontuario, censo.nm_paciente, censo.in_sexo,
                    censo.nr_idade, censo.dt_nascimento, censo.vinculo, censo.nm_especialidade,
                    censo.nm_medico, censo.cod_CID, censo.descricaoCID, censo.tempo, censo.nm_origem,
                    censo.dt_ultimo_evento_data, censo.dt_ultimo_evento_hora, censo.nr_convenio
                );
            }

            return dt;
        }


        public static void ExportDataTableToExcel(DataTable dataTable, string filePath)
        {
            // Verifica se o DataTable é nulo ou vazio
            if (dataTable == null || dataTable.Rows.Count == 0)
            {
                throw new Exception("O DataTable está vazio ou nulo.");
            }

            // Define o contexto de licença do EPPlus (necessário para versões >= 5.0)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Cria um novo arquivo Excel
            using (var package = new ExcelPackage())
            {
                // Adiciona uma planilha ao arquivo Excel
                var worksheet = package.Workbook.Worksheets.Add("Dados");

                // Escreve os cabeçalhos das colunas
                for (int col = 0; col < dataTable.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                }

                // Escreve os dados das linhas
                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dataTable.Rows[row][col];
                    }
                }
                try
                {
                    // Write column headings
                    for (int i = 0; i < dataTable.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
                    }

                    // Write rows
                    for (int row = 0; row < dataTable.Rows.Count; row++)
                    {
                        for (int col = 0; col < dataTable.Columns.Count; col++)
                        {
                            if (col == 2 || col == 8 || col == 16) // Columns with dates
                            {
                                DateTime dateValue = SafeConvertToDateTime(dataTable.Rows[row][col], DateTime.MinValue);
                                if (dateValue.Equals(DateTime.MinValue))
                                {
                                    worksheet.Cells[row + 2, col + 1].Value = "";
                                }
                                else
                                {
                                    worksheet.Cells[row + 2, col + 1].Value =   dateValue;
                                    worksheet.Cells[row + 2, col + 1].Style.Numberformat.Format = "dd/MM/yyyy"; // Formato de data abreviada

                                }
                            }
                            else
                            {
                                worksheet.Cells[row + 2, col + 1].Value =  dataTable.Rows[row][col];
                            }
                        }
                    }

                    // Salva o arquivo Excel no caminho especificado
                    package.SaveAs(new System.IO.FileInfo(filePath));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Arquivo Excel salvo com sucesso em: {filePath}");
                }

            }
        }
        //public static void ExportarParaExcel(DataTable dataCenso, string filePath)
        //{
        //    if (dataCenso == null || dataCenso.Rows.Count == 0)
        //        throw new Exception("Nenhum dado disponível para exportação.");

        //    Excel.Application excelApp = new Excel.Application { Visible = false };
        //    Excel.Workbook wb = excelApp.Workbooks.Add(Excel.XlSheetType.xlWorksheet);
        //        // Acessa a primeira planilha diretamente
        //        Excel.Worksheet ws = wb.Sheets[1] as Excel.Worksheet;

        //        try
        //    {
        //        // Write column headings
        //        for (int i = 0; i < dataCenso.Columns.Count; i++)
        //        {
        //            ws.Cells[1, i + 1] = dataCenso.Columns[i].ColumnName;
        //        }

        //        // Write rows
        //        for (int i = 0; i < dataCenso.Rows.Count; i++)
        //        {
        //            for (int j = 0; j < dataCenso.Columns.Count; j++)
        //            {
        //                if (j == 2 || j == 8 || j == 16) // Columns with dates
        //                {
        //                    DateTime dateValue = SafeConvertToDateTime(dataCenso.Rows[i][j], DateTime.MinValue);
        //                        if (dateValue.Equals(DateTime.MinValue))
        //                        {
        //                            ws.Cells[i + 2, j + 1] = "";
        //                        }
        //                        else
        //                        {
        //                            ws.Cells[i + 2, j + 1] = dateValue;
        //                        }
        //                }
        //                else
        //                {
        //                    ws.Cells[i + 2, j + 1] = dataCenso.Rows[i][j];
        //                }
        //            }
        //        }
        //            if (ws == null)
        //            {
        //                throw new Exception("A planilha não foi criada corretamente.");
        //            }
        //            // Obtém o intervalo de células usadas
        //            Excel.Range usedRange = ws.UsedRange;

        //            // Verifica se o intervalo é nulo
        //            if (usedRange == null)
        //            {
        //                Console.WriteLine("O intervalo de células usadas é nulo.");
        //                return;
        //            }

        //            // Verifica se o intervalo está vazio
        //            if (usedRange.Rows.Count == 0 || usedRange.Columns.Count == 0)
        //            {
        //                Console.WriteLine("A planilha está vazia.");
        //                return;
        //            }

        //            // Define the range using the correct syntax
        //            Excel.Range dataRange = ws.get_Range(
        //                ws.Cells[1, 1], // Start cell (top-left corner)
        //                ws.Cells[dataCenso.Rows.Count , dataCenso.Columns.Count] // End cell (bottom-right corner)
        //            );
        //            Excel.ListObject excelTable = ws.ListObjects.AddEx(
        //                SourceType: Excel.XlListObjectSourceType.xlSrcRange,
        //                Source: dataRange,
        //                XlListObjectHasHeaders: Excel.XlYesNoGuess.xlYes
        //            );
        //            excelTable.Name = "CensoTable"; // Name the table
        //                                            // Save the file
        //            if (!string.IsNullOrEmpty(filePath))
        //        {
        //            try
        //            {
        //                ws.Name = "Censo";
        //                wb.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, Excel.XlSaveAsAccessMode.xlNoChange,
        //                          Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
        //                Console.WriteLine("Excel file saved!");
        //            }
        //            catch (Exception ex)
        //            {
        //                throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n" + ex.Message);
        //            }
        //        }
        //        else
        //        {
        //            excelApp.Visible = true;
        //            Console.WriteLine("A tabela está vazia, impossível formatar.");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Erro ao salvar o arquivo Excel: {ex.Message}");
        //    }
        //    finally
        //    {
        //        wb.Close(false);
        //        excelApp.Quit();

        //        // Release COM objects
        //        Marshal.ReleaseComObject(ws);
        //        Marshal.ReleaseComObject(wb);
        //        Marshal.ReleaseComObject(excelApp);
        //    }
        //}



        private static string TratarTempo(string tempo)
        {
            return tempo?.Replace("days", " ")
                         .Replace("day", " ")
                         .Replace("00:00:00", "0")
                         ?? " ";
        }

    }
        public class Censo
        {
            public string nr_quarto { get; set; }
            public string nm_unidade_funcional { get; set; }
            public string dt_internacao_data { get; set; }
            public string dt_internacao_hora { get; set; }
            public string cd_prontuario { get; set; }
            public string nm_paciente { get; set; }
            public string in_sexo { get; set; }
            public string nr_idade { get; set; }
            public string dt_nascimento { get; set; }
            public string vinculo { get; set; }
            public string nm_especialidade { get; set; }
            public string nm_medico { get; set; }
            public string cod_CID { get; set; }
            public string descricaoCID { get; set; }
            public string tempo { get; set; }
            public string nm_origem { get; set; }
            public string dt_ultimo_evento_data { get; set; }
            public string dt_ultimo_evento_hora { get; set; }
            public string nr_convenio { get; set; }
        }
    }





    