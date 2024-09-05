using Azure.Storage.Blobs; // Para interactuar con Azure Blob Storage
using Microsoft.AspNetCore.Http; // Para manejar solicitudes HTTP
using Microsoft.AspNetCore.Mvc; // Para manejar respuestas HTTP
using Microsoft.Azure.WebJobs.Extensions.Http; // Para funciones HTTP en Azure Functions
using Microsoft.Azure.WebJobs; // Para definir Azure Functions
using Microsoft.Extensions.Logging; // Para registrar logs
using OfficeOpenXml; // Para manejar archivos Excel
using PdfSharp.Drawing; // Para manejar gráficos en PDF
using PdfSharp.Pdf; // Para generar archivos PDF
using System.IO; // Para manejo de flujos de datos
using System.Threading.Tasks; // Para manejar operaciones asíncronas
using System; // Para funcionalidades básicas de .NET
using Microsoft.Data.SqlClient; // Para interactuar con SQL Server
using PdfSharp.Charting; // Para crear gráficos y tablas
using PdfSharp.Drawing.Layout; // Para manejo avanzado de texto
using Azure.Identity; // Para autenticación con Azure
using Azure.Security.KeyVault.Secrets; // Para acceder a secretos en Key Vault

public static class ReporteVentasFunction
{
    [FunctionName("GenerarReporteVentas")] // Define el nombre de la función Azure
    public static async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, // Define el desencadenador HTTP
        ILogger log) // Logger para registrar información
    {
        log.LogInformation("Iniciando generación del reporte de ventas mensuales.");

        try
        {
            // Obtener la URL de Key Vault desde las variables de entorno
            string keyVaultUrl = Environment.GetEnvironmentVariable("KeyVaultUrl");

            if (string.IsNullOrEmpty(keyVaultUrl))
            {
                throw new InvalidOperationException("La URL de Key Vault no está configurada o es nula.");
            }

            // Autenticar y conectarse a Key Vault
            var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential());

            // Obtener las cadenas de conexión desde Key Vault
            KeyVaultSecret sqlConnectionStringSecret = await client.GetSecretAsync("SQLConnectionString");
            string sqlConnectionString = sqlConnectionStringSecret.Value;

            KeyVaultSecret blobConnectionStringSecret = await client.GetSecretAsync("BlobStorageConnectionString");
            string blobConnectionString = blobConnectionStringSecret.Value;

            log.LogInformation("Secretos obtenidos exitosamente desde Key Vault.");

            // Configurar la licencia de EPPlus para uso no comercial
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (SqlConnection conn = new SqlConnection(sqlConnectionString))
            {
                // Establecer conexión con la base de datos
                await conn.OpenAsync();
                log.LogInformation("Conexión a la base de datos establecida con éxito.");

                // Nueva consulta SQL para obtener ventas mensuales agrupadas por año y mes
                string query = @"
                    SELECT 
                        YEAR(FechaVenta) AS Año,
                        MONTH(FechaVenta) AS Mes,
                        SUM(Cantidad * Precio) AS TotalVentas
                    FROM Productos
                    WHERE FechaVenta IS NOT NULL
                    GROUP BY YEAR(FechaVenta), MONTH(FechaVenta)
                    ORDER BY Año, Mes";

                SqlCommand cmd = new SqlCommand(query, conn);

                using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                {
                    log.LogInformation("Consulta SQL ejecutada exitosamente.");

                    // Crear un nuevo paquete de Excel
                    using (var excelPackage = new ExcelPackage())
                    {
                        // Crear una hoja de trabajo en Excel y agregar encabezados
                        var ws = excelPackage.Workbook.Worksheets.Add("Reporte de Ventas Mensuales");
                        ws.Cells["A1"].Value = "Año";
                        ws.Cells["B1"].Value = "Mes";
                        ws.Cells["C1"].Value = "Total Ventas";

                        int row = 2; // Comienza en la segunda fila para llenar datos
                        while (await reader.ReadAsync())
                        {
                            // Rellenar la hoja de Excel con los datos de ventas mensuales
                            ws.Cells[row, 1].Value = reader["Año"];
                            ws.Cells[row, 2].Value = reader["Mes"];
                            ws.Cells[row, 3].Value = reader["TotalVentas"];
                            row++;
                        }

                        // Guardar el contenido de Excel en un flujo de memoria
                        using (var excelStream = new MemoryStream())
                        {
                            excelPackage.SaveAs(excelStream);
                            excelStream.Position = 0;

                            // Crear un nuevo documento PDF
                            using (var pdfStream = new MemoryStream())
                            {
                                var pdfDoc = new PdfDocument();
                                var pdfPage = pdfDoc.AddPage();
                                var gfx = XGraphics.FromPdfPage(pdfPage);

                                // Definir fuentes y colores
                                var fontBold = new XFont("Verdana", 12, XFontStyleEx.Bold);
                                var fontRegular = new XFont("Verdana", 10, XFontStyleEx.Regular);
                                var tableHeaderBrush = XBrushes.DarkBlue;
                                var tableCellBrush = XBrushes.LightGray;

                                // Agregar título al PDF
                                gfx.DrawString("Reporte de Ventas Mensuales", fontBold, XBrushes.Black,
                                    new XRect(0, 0, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopCenter);

                                // Descripción del reporte en el PDF
                                double yPoint = 40;
                                gfx.DrawString("Este reporte muestra el total de ventas agrupadas por mes.", fontRegular, XBrushes.Black,
                                    new XRect(20, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft);

                                // Dibujar tabla
                                yPoint += 40;

                                // Definir posiciones iniciales y tamaños
                                double xPos = 20;
                                double cellHeight = 20;
                                double cellWidthYear = 80;
                                double cellWidthMonth = 80;
                                double cellWidthTotal = 100;

                                // Dibujar encabezados de la tabla
                                gfx.DrawRectangle(tableHeaderBrush, xPos, yPoint, cellWidthYear, cellHeight);
                                gfx.DrawString("Año", fontBold, XBrushes.White, new XRect(xPos, yPoint, cellWidthYear, cellHeight), XStringFormats.Center);

                                gfx.DrawRectangle(tableHeaderBrush, xPos + cellWidthYear, yPoint, cellWidthMonth, cellHeight);
                                gfx.DrawString("Mes", fontBold, XBrushes.White, new XRect(xPos + cellWidthYear, yPoint, cellWidthMonth, cellHeight), XStringFormats.Center);

                                gfx.DrawRectangle(tableHeaderBrush, xPos + cellWidthYear + cellWidthMonth, yPoint, cellWidthTotal, cellHeight);
                                gfx.DrawString("Total Ventas", fontBold, XBrushes.White, new XRect(xPos + cellWidthYear + cellWidthMonth, yPoint, cellWidthTotal, cellHeight), XStringFormats.Center);

                                yPoint += cellHeight;

                                // Reposicionar el lector de datos SQL para leer nuevamente los resultados para el PDF
                                reader.Close();
                                cmd.CommandText = query;
                                using (SqlDataReader pdfReader = await cmd.ExecuteReaderAsync())
                                {
                                    // Rellenar el contenido de la tabla con los datos de ventas mensuales
                                    while (await pdfReader.ReadAsync())
                                    {
                                        gfx.DrawRectangle(tableCellBrush, xPos, yPoint, cellWidthYear, cellHeight);
                                        gfx.DrawString(pdfReader["Año"].ToString(), fontRegular, XBrushes.Black, new XRect(xPos, yPoint, cellWidthYear, cellHeight), XStringFormats.Center);

                                        gfx.DrawRectangle(tableCellBrush, xPos + cellWidthYear, yPoint, cellWidthMonth, cellHeight);
                                        gfx.DrawString(pdfReader["Mes"].ToString(), fontRegular, XBrushes.Black, new XRect(xPos + cellWidthYear, yPoint, cellWidthMonth, cellHeight), XStringFormats.Center);

                                        gfx.DrawRectangle(tableCellBrush, xPos + cellWidthYear + cellWidthMonth, yPoint, cellWidthTotal, cellHeight);
                                        gfx.DrawString(pdfReader["TotalVentas"].ToString(), fontRegular, XBrushes.Black, new XRect(xPos + cellWidthYear + cellWidthMonth, yPoint, cellWidthTotal, cellHeight), XStringFormats.Center);

                                        yPoint += cellHeight;
                                    }
                                }

                                // Guardar el contenido del PDF en un flujo de memoria
                                pdfDoc.Save(pdfStream);
                                pdfStream.Position = 0;

                                // Conectar a Azure Blob Storage
                                BlobServiceClient blobServiceClient = new BlobServiceClient(blobConnectionString);
                                BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient("reportes");

                                // Crear el contenedor en Blob Storage si no existe
                                await containerClient.CreateIfNotExistsAsync();
                                log.LogInformation("Contenedor de Blob Storage verificado o creado.");

                                // Generar nombres únicos para los archivos usando un timestamp
                                string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                                string excelBlobName = $"ReporteVentasMensuales_{timestamp}.xlsx";
                                BlobClient excelBlob = containerClient.GetBlobClient(excelBlobName);
                                excelStream.Position = 0;
                                // Subir el archivo Excel a Blob Storage
                                await excelBlob.UploadAsync(excelStream, overwrite: true);
                                log.LogInformation($"Archivo Excel subido exitosamente como {excelBlobName}.");

                                string pdfBlobName = $"ReporteVentasMensuales_{timestamp}.pdf";
                                BlobClient pdfBlob = containerClient.GetBlobClient(pdfBlobName);
                                pdfStream.Position = 0;
                                // Subir el archivo PDF a Blob Storage
                                await pdfBlob.UploadAsync(pdfStream, overwrite: true);
                                log.LogInformation($"Archivo PDF subido exitosamente como {pdfBlobName}.");

                                // Devolver una respuesta HTTP indicando éxito
                                return new OkObjectResult($"Archivos subidos exitosamente: {excelBlobName}, {pdfBlobName}");
                            }
                        }
                    }
                }
            }
        }
        catch (InvalidOperationException ex)
        {
            log.LogError($"Error de configuración: {ex.Message}");
            return new StatusCodeResult(StatusCodes.Status500InternalServerError); // Error 500 en caso de problema de configuración
        }
        catch (SqlException ex)
        {
            log.LogError($"Error al ejecutar la consulta SQL: {ex.Message}");
            return new StatusCodeResult(StatusCodes.Status500InternalServerError); // Error 500 en caso de problema con la consulta SQL
        }
        catch (Azure.RequestFailedException ex)
        {
            log.LogError($"Error al interactuar con Azure Blob Storage: {ex.Message}");
            return new StatusCodeResult(StatusCodes.Status500InternalServerError); // Error 500 en caso de problema con Azure Blob Storage
        }
        catch (Exception ex)
        {
            log.LogError($"Error general al generar el reporte: {ex.Message}");
            return new StatusCodeResult(StatusCodes.Status500InternalServerError); // Error 500 en caso de cualquier otro problema
        }
    }
}
