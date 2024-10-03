using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.SqlClient;
using System.Globalization;

class Program
{
    static async Task Main(string[] args)
    {
        string connectionString = "Tu Conexion";

        try
        {
            // Ruta del archivo de Excel
            string? filePath;

            while (true)
            {
                Console.Write("Ingrese la ruta del archivo de Excel: ");
                filePath = Console.ReadLine();

                if (!String.IsNullOrEmpty(filePath))
                {
                    break;
                }
                else
                {
                    Console.WriteLine("La ruta del archivo de Excel no puede ser vacía. Intente nuevamente.");
                }
            }


            // Lista para almacenar los resultados
            List<Result> results = new List<Result>();

            #region Leer el archivo de Excel

            using (var spreadsheetDocument = SpreadsheetDocument.Open(filePath, true))
            {
                foreach (var sheet in spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>())
                {
                    // Obtener la hoja de trabajo
                    var worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id);
                    var worksheet = worksheetPart.Worksheet;

                    // Extraer valores EUR (fila 11, columnas F y G)
                    var eurValueF = decimal.Parse(GetCellValue(worksheet, 11, 6), NumberStyles.Any, CultureInfo.InvariantCulture);
                    var eurValueG = decimal.Parse(GetCellValue(worksheet, 11, 7), NumberStyles.Any, CultureInfo.InvariantCulture);


                    // Extraer valores USD (fila 15, columnas F y G)
                    var usdValueF = decimal.Parse(GetCellValue(worksheet, 15, 6), NumberStyles.Any, CultureInfo.InvariantCulture);
                    var usdValueG = decimal.Parse(GetCellValue(worksheet, 15, 7), NumberStyles.Any, CultureInfo.InvariantCulture);

                    var sheetName = sheet.Name;
                    var fecha = DateTime.ParseExact(sheetName, "ddMMyyyy", null);
                    

                    results.Add(new Result
                        {
                            Sheet = sheet.Name,
                            EUR_F = eurValueF,
                            EUR_G = eurValueG,
                            USD_F = usdValueF,
                            USD_G = usdValueG,
                            FechaOperacion = fecha,
                    });
                    
                }
            }

            #endregion

            #region Crear archivo de texto

            string outputFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "results.txt");
            using (StreamWriter writer = new StreamWriter(outputFile))
            {
                foreach (var result in results)
                {
                    writer.WriteLine($"Hoja: {result.Sheet}");
                    writer.WriteLine($"EUR C: {result.EUR_F}");
                    writer.WriteLine($"EUR V: {result.EUR_G}");
                    writer.WriteLine($"USD C: {result.USD_F}");
                    writer.WriteLine($"USD V: {result.USD_G}");
                    writer.WriteLine($"Fecha: {result.FechaOperacion}");
                    writer.WriteLine();
                }
            }

            Console.WriteLine($"Resultados guardados en {outputFile}");
            #endregion

             
            #region Generar script de inserción 

            #region Generar script de inserción Administrativo
            string scriptFileAdministrativo = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "scriptAdministrativo.sql");
            using (StreamWriter writer = new StreamWriter(scriptFileAdministrativo))
            {

                writer.WriteLine("IF exists(select* from saMoneda where co_mone = N'USD   ' )");
                writer.WriteLine("BEGIN");

                foreach (var result in results)
                {
                    var tasaUSD = new { moneda = "USD", tasaC = result.USD_F, tasaV = result.USD_G };
                    writer.WriteLine($"IF NOT EXISTS (select co_mone from satasa where co_mone = N'{tasaUSD.moneda}   ' and fecha = CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime) ) INSERT [dbo].[satasa] ([co_mone], [fecha], [tasa_c], [tasa_v], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [co_sucu_in], [fe_us_in], [co_us_mo], [co_sucu_mo], [fe_us_mo], [revisado], [trasnfe])  VALUES (N'{tasaUSD.moneda}   ', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime), CAST({tasaUSD.tasaC.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), CAST({tasaUSD.tasaV.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, N'PROFIT', NULL, CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), N'PROFIT', NULL, CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), NULL, NULL)");
                }
                writer.WriteLine("END");

                writer.WriteLine(" ");

                writer.WriteLine("IF exists(select* from saMoneda where co_mone = N'EUR   ' )");
                writer.WriteLine("BEGIN");
                foreach (var result in results)
                {
                    var tasaEUR = new { moneda = "EUR", tasaC = result.EUR_F, tasaV = result.EUR_G };
                    writer.WriteLine($"IF NOT EXISTS (select co_mone from satasa where co_mone = N'{tasaEUR.moneda}   ' and fecha = CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime) ) INSERT [dbo].[satasa] ([co_mone], [fecha], [tasa_c], [tasa_v], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [co_sucu_in], [fe_us_in], [co_us_mo], [co_sucu_mo], [fe_us_mo], [revisado], [trasnfe])  VALUES (N'{tasaEUR.moneda}   ', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime), CAST({tasaEUR.tasaC.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), CAST({tasaEUR.tasaV.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, N'PROFIT', NULL, CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), N'PROFIT', NULL, CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), NULL, NULL)");
                }
                writer.WriteLine("END");
            }
            #endregion

            #region Generar script de inserción Nomina
            string scriptFileNomina = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "scriptNomina.sql");
            using (StreamWriter writer = new StreamWriter(scriptFileNomina))
            {

                writer.WriteLine("IF exists(select* from snMoneda where co_mone = N'USD   ' )");
                writer.WriteLine("BEGIN");

                foreach (var result in results)
                {
                    var tasaUSD = new { moneda = "USD", tasaC = result.USD_F, tasaV = result.USD_G };
                    writer.WriteLine($"IF NOT EXISTS (select co_mone from sntasa where co_mone = N'{tasaUSD.moneda}   ' and fecha = CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime) ) INSERT [dbo].[sntasa] ([co_mone], [fecha], [tasa_c], [tasa_v], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [revisado], [trasnfe])  VALUES (N'{tasaUSD.moneda}   ', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime), CAST({tasaUSD.tasaC.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), CAST({tasaUSD.tasaV.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, N'PROFIT', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), N'PROFIT', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), NULL, NULL)");
                }
                writer.WriteLine("END");

                writer.WriteLine(" ");

                writer.WriteLine("IF exists(select* from snMoneda where co_mone = N'EUR   ' )");
                writer.WriteLine("BEGIN");
                foreach (var result in results)
                {
                    var tasaEUR = new { moneda = "EUR", tasaC = result.EUR_F, tasaV = result.EUR_G };
                    writer.WriteLine($"IF NOT EXISTS (select co_mone from sntasa where co_mone = N'{tasaEUR.moneda}   ' and fecha = CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime) ) INSERT [dbo].[sntasa] ([co_mone], [fecha], [tasa_c], [tasa_v], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [revisado], [trasnfe])  VALUES (N'{tasaEUR.moneda}   ', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime), CAST({tasaEUR.tasaC.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), CAST({tasaEUR.tasaV.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, N'PROFIT', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), N'PROFIT', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), NULL, NULL)");
                }
                writer.WriteLine("END");
            }
            #endregion

            #region Generar script de inserción Contabilidad
            string scriptFileContabilidad = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "scriptContabilidad.sql");
            using (StreamWriter writer = new StreamWriter(scriptFileContabilidad))
            {

                writer.WriteLine("IF exists(select* from scMoneda where co_mone = N'USD   ' )");
                writer.WriteLine("BEGIN");

                foreach (var result in results)
                {
                    var tasaUSD = new { moneda = "USD", tasaC = result.USD_F, tasaV = result.USD_G };
                    writer.WriteLine($"IF NOT EXISTS (select co_mone from sctasa where co_mone = N'{tasaUSD.moneda}   ' and fecha = CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime) ) INSERT [dbo].[sctasa] ([co_mone], [fecha], [tasa_c], [tasa_v], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [revisado], [trasnfe])  VALUES (N'{tasaUSD.moneda}   ', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime), CAST({tasaUSD.tasaC.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), CAST({tasaUSD.tasaV.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, N'PROFIT', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), N'PROFIT', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), NULL, NULL)");
                }
                writer.WriteLine("END");

                writer.WriteLine(" ");

                writer.WriteLine("IF exists(select* from scMoneda where co_mone = N'EUR   ' )");
                writer.WriteLine("BEGIN");
                foreach (var result in results)
                {
                    var tasaEUR = new { moneda = "EUR", tasaC = result.EUR_F, tasaV = result.EUR_G };
                    writer.WriteLine($"IF NOT EXISTS (select co_mone from sctasa where co_mone = N'{tasaEUR.moneda}   ' and fecha = CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime) ) INSERT [dbo].[sctasa] ([co_mone], [fecha], [tasa_c], [tasa_v], [campo1], [campo2], [campo3], [campo4], [campo5], [campo6], [campo7], [campo8], [co_us_in], [fe_us_in], [co_us_mo], [fe_us_mo], [revisado], [trasnfe])  VALUES (N'{tasaEUR.moneda}   ', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS SmallDateTime), CAST({tasaEUR.tasaC.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), CAST({tasaEUR.tasaV.ToString(CultureInfo.InvariantCulture)} AS Decimal(21, 8)), NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, N'PROFIT', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), N'PROFIT', CAST(N'{result.FechaOperacion.ToString("yyyy-MM-dd HH:mm:ss")}' AS DateTime), NULL, NULL)");
                }
                writer.WriteLine("END");
            }
            #endregion

            Console.WriteLine($"Script generado en Downloads/Descargas");
            #endregion   

            await GuardarTasa(results, connectionString);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
        }

    }



    static async Task GuardarTasa(List<Result> results, string connectionString)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            string queryTasas = "INSERT INTO tasa (co_mone, fecha, tasa_c, tasa_v) " +
                               "VALUES (@moneda, @fecha, @tasa_c, @tasa_v) ";

            foreach (var result in results)
            {
                var tasas = new[]
                {
                new { moneda = "USD", tasaC = result.USD_F, tasaV = result.USD_G },
                new { moneda = "EUR", tasaC = result.EUR_F, tasaV = result.EUR_G }
            };

                foreach (var tasa in tasas)
                {
                    // Verificar si ya existe un registro con la misma fecha de operación y nombre de hoja
                    string queryExists = "SELECT COUNT(*) FROM tasa WHERE co_mone = @moneda AND fecha = @fecha";
                    using (SqlCommand commandExists = new SqlCommand(queryExists, connection))
                    {
                        commandExists.Parameters.AddWithValue("@moneda", tasa.moneda);
                        commandExists.Parameters.AddWithValue("@fecha", result.FechaOperacion);

                        int count = (int)commandExists.ExecuteScalar();

                        if (count > 0)
                        {
                            // Ya existe un registro con la misma fecha de operación y nombre de hoja, actualizar registro existente
                            string queryUpdate = "UPDATE tasa SET tasa_c = @tasa_c, tasa_v = @tasa_v WHERE co_mone = @moneda AND fecha = @fecha";
                            using (SqlCommand command = new SqlCommand(queryUpdate, connection))
                            {
                                command.Parameters.AddWithValue("@moneda", tasa.moneda);
                                command.Parameters.AddWithValue("@fecha", result.FechaOperacion);
                                command.Parameters.AddWithValue("@tasa_c", tasa.tasaC);
                                command.Parameters.AddWithValue("@tasa_v", tasa.tasaV);
                                command.ExecuteNonQuery();
                            }
                        }
                        else
                        {
                            // No existe un registro con la misma fecha de operación y nombre de hoja, insertar nuevo registro
                            using (SqlCommand command = new SqlCommand(queryTasas, connection))
                            {
                                command.Parameters.AddWithValue("@moneda", tasa.moneda);
                                command.Parameters.AddWithValue("@fecha", result.FechaOperacion);
                                command.Parameters.AddWithValue("@tasa_c", tasa.tasaC);
                                command.Parameters.AddWithValue("@tasa_v", tasa.tasaV);
                                command.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
        }
    }

    static string GetCellValue(Worksheet worksheet, int row, int column)
    {
        var cellReference = $"{GetColumnLetter(column)}{row}";
        var cell = worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellReference);
        if (cell != null)
        {
            return cell.InnerText;
        }
        return string.Empty;
    }

    static string GetColumnLetter(int column)
    {
        var columnLetter = string.Empty;
        while (column > 0)
        {
            column--;
            columnLetter = ((char)('A' + (column % 26))) + columnLetter;
            column /= 26;
        }
        return columnLetter;
    }

}


public class Result
{
    public string Sheet { get; set; }
    public decimal EUR_F { get; set; }
    public decimal EUR_G { get; set; }
    public decimal USD_F { get; set; }
    public decimal USD_G { get; set; }
    public DateTime FechaOperacion { get; set; }
}
