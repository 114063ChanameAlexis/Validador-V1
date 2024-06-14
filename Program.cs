using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Validador.Clases;

namespace Validador
{
    internal class Program
    {
        static string filePath = "C:\\Users\\alexis.chaname\\Desktop\\Wuala\\Prueba.xlsx";
        static string sheetName = "Test";
        static List<Publicacion> itemsList = new List<Publicacion>();
        static async Task Main(string[] args)
        {
            List<string> mlas = CargarMlasDesdeExcel(filePath, sheetName);
            Console.WriteLine($"\nCantidad de Publicaciones: {mlas.Count}");

            int segundosCalculados = mlas.Count * 4;
            TimeSpan tiempo = TimeSpan.FromSeconds(segundosCalculados);
            Console.WriteLine($"Tiempo estimado de consulta: {tiempo.Hours} horas - {tiempo.Minutes} minutos - {tiempo.Seconds} segundos");

            while (true)
            {
                for (int i = 0; i < mlas.Count; i++)
                {
                    string mla = mlas[i];
                    Console.WriteLine($" ---- \nConsultando publicacion {i + 1}/{mlas.Count}: {mla}");
                    await ConsultarApiAsync(mla);
                    await Task.Delay(2000);
                }
                GuardarEnExcel();

                Console.WriteLine("Presiona Enter para consultar los items nuevamente o escribe 'exit' para salir.");
                string userInput = Console.ReadLine();

                if (userInput.Equals("exit", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }
            }

        }
        static List<string> CargarMlasDesdeExcel(string filePath, string sheetName)
        {
            List<string> mlas = new List<string>();

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {
                    // Obtener la hoja de trabajo por su nombre
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[sheetName];

                    // Obtener el número total de filas en la hoja de trabajo
                    int rowCount = worksheet.Dimension.Rows;

                    // Número de la columna que contiene los códigos MLA (por ejemplo, columna A)
                    int mlaColumnNumber = 1;

                    // Número de la columna que contiene los valores a convertir a decimal (por ejemplo, columna B)
                    //int decimalColumnNumber = 2;

                    // Recorrer las filas para obtener los valores de la columna MLA y convertir la columna a decimal
                    for (int row = 2; row <= rowCount; row++) // Comenzamos desde la fila 2, asumiendo que la fila 1 son los encabezados
                    {
                        // Obtener el valor de la celda en la columna MLA
                        string mlaValue = worksheet.Cells[row, mlaColumnNumber].Value?.ToString();

                        // Agregar el valor a la lista de MLAs
                        if (!string.IsNullOrEmpty(mlaValue))
                        {
                            mlas.Add(mlaValue);
                        }

                        // Convertir el valor de la celda de la columna a decimal si es posible
                        //ExcelRange cell = worksheet.Cells[row, decimalColumnNumber];
                        //if (cell.Value != null)
                        //{
                        //    string cellValueAsString = cell.Value.ToString();
                        //    if (double.TryParse(cellValueAsString, out double cellValueAsDouble))
                        //    {
                        //        cell.Style.Numberformat.Format = "0.00";
                        //        // Si la conversión es exitosa, establece el nuevo valor de la celda como decimal
                        //        cell.Value = cellValueAsDouble;

                        //    }
                        //}
                    }
                    // Guardar los cambios en el archivo Excel
                    excelPackage.Save();
                }

                Console.WriteLine("Valores de MLA cargados desde el archivo Excel:");
                foreach (string mla in mlas)
                {
                    Console.WriteLine(mla);
                }

                return mlas;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al leer el archivo Excel: {ex.Message}");
                return new List<string>();
            }
        }
        static async Task ConsultarApiAsync(string mla)
        {
            string apiUrl = $"https://api.mercadolibre.com/items/{mla}";

            using (HttpClient httpClient = new HttpClient())
            {
                try
                {
                    HttpResponseMessage response = await httpClient.GetAsync(apiUrl);

                    if (response.IsSuccessStatusCode)
                    {
                        string contenidoJson = await response.Content.ReadAsStringAsync();
                        Publicacion mercadoLibreItem = Newtonsoft.Json.JsonConvert.DeserializeObject<Publicacion>(contenidoJson);
                        itemsList.Add(mercadoLibreItem);
                        //if (mercadoLibreItem.Shipping == null)
                        //{
                        //    mercadoLibreItem.Shipping = new Shipping();
                        //    mercadoLibreItem.Shipping.logistic_type = "No tiene";
                        //}

                        //Alexis 13/06////////////if (mercadoLibreItem.Variations != null)
                        ////////////{
                        ////////////    foreach (Variacion var in mercadoLibreItem.Variations)
                        ////////////    {
                        ////////////        string SKU = await ConsultarVariation(mercadoLibreItem.Id, var.Id);
                        ////////////        await Task.Delay(2000);
                        ////////////        var.SKU = SKU;
                        ////////////        Console.WriteLine($"ID de variantes: {var.Id} | SKU: {var.SKU}");
                        ////////////    }
                        ////////////}
                        ////////////Console.WriteLine($"-> Precio: {mercadoLibreItem.Price} | Precio base: {mercadoLibreItem.Base_price} | Status: {mercadoLibreItem.Status} | Es de catalogo: {mercadoLibreItem.Catalog_listing} | Logistica: {mercadoLibreItem.Shipping.logistic_type}");
                    }
                    else
                    {
                        Console.WriteLine($"Error en la solicitud para el item {mla}: {response.StatusCode} - {response.ReasonPhrase}");
                        Publicacion mercadoLibreItem = new Publicacion();
                        itemsList.Add(mercadoLibreItem);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error en la solicitud para el item {mla}: {ex.Message}");
                    Publicacion mercadoLibreItem = new Publicacion();
                    itemsList.Add(mercadoLibreItem);
                }
            }
        }
        static async Task <string> ConsultarVariation(string publicacion, string variacion)
        {
            string urlVariation = $"https://api.mercadolibre.com/items/{publicacion}/variations/{variacion}";

            using (HttpClient httpClient = new HttpClient())
            {
                try
                {
                    HttpResponseMessage response = await httpClient.GetAsync(urlVariation);

                    if (response.IsSuccessStatusCode)
                    {
                        string contenidoJson = await response.Content.ReadAsStringAsync();
                        Variacion variationJson = Newtonsoft.Json.JsonConvert.DeserializeObject<Variacion>(contenidoJson);
                        string sku = variationJson.Attributes.FirstOrDefault(attr => attr.Id == "SELLER_SKU")?.Value_Name;
                        return sku;
                    }
                    else
                    {
                        Console.WriteLine($"Error en la solicitud. Código de estado: {response.StatusCode}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                }
                finally
                {
                    httpClient.Dispose();
                }
            }
            return null;
        }
        static void GuardarEnExcel()
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                string folderPath = "C:\\Users\\alexis.chaname\\Desktop\\Wuala";
                string filePath = Path.Combine(folderPath, "Prueba.xlsx");
                FileInfo existingFile = new FileInfo(filePath);

                using (ExcelPackage excelPackage = existingFile.Exists ? new ExcelPackage(existingFile) : new ExcelPackage())
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Test") ?? excelPackage.Workbook.Worksheets.Add("MercadoLibreItems");

                    worksheet.Cells[1, 5].Value = "Es Full";
                    worksheet.Cells[1, 6].Value = "MLA";
                    worksheet.Cells[1, 7].Value = "Precio";
                    worksheet.Cells[1, 8].Value = "Estado";
                    worksheet.Cells[1, 9].Value = "Catalogo";
                    //worksheet.Cells[1, 10].Value = "Condicion";

                    int maxVariationsCount = itemsList.Where(publicacion => publicacion != null).Max(publicacion => publicacion.Variations?.Count ?? 0);

                    for (int i = 1; i <= maxVariationsCount; i++)
                    {
                        worksheet.Cells[1, 9 + 2 * i].Value = $"ID Variante {i}";
                        worksheet.Cells[1, 10 + 2 * i].Value = $"SKU {i}";
                    }

                    int row = 2;

                    foreach (Publicacion item in itemsList)
                    {
                        worksheet.Cells[row, 5].Value = item.Shipping.logistic_type;
                        worksheet.Cells[row, 6].Value = item.Id;
                        worksheet.Cells[row, 7].Value = item.Price;
                        worksheet.Cells[row, 8].Value = item.Status;
                        worksheet.Cells[row, 9].Value = item.Catalog_listing;

                        int cellVarId = 11;
                        int cellVarSKU = 12;

                        if (item.Variations != null && item.Variations.Count > 0)
                        //if (item.Variations.Count > 0 && item.Variations != null)
                        {
                            if (item.Variations[0].SKU == null)
                            {
                                item.Variations[0].SKU = "no tiene SKU";
                            }
                            foreach (Variacion variation in item.Variations)
                            {
                                worksheet.Cells[row, cellVarId].Value = variation.Id;
                                worksheet.Cells[row, cellVarSKU].Value = variation.SKU;
                                cellVarId += 2;
                                cellVarSKU += 2;
                            }
                            //decimal valorCeldaFila2 = Convert.ToDecimal(worksheet.Cells[row, 2].Value);
                            if (item.Variations[0].SKU.Equals(worksheet.Cells[row, 3].Value))
                            {
                                worksheet.Cells[row, 4].Value = "Correcto";
                            }
                            else if (item.Variations[0].SKU.Equals("no tiene SKU"))
                            {
                                worksheet.Cells[row, 4].Value = "Cargar SKU";
                            }
                            else
                            {
                                worksheet.Cells[row, 4].Value = "Revisar SKU";
                            }
                        }
                        row++;
                    }
                        excelPackage.SaveAs(existingFile);
                }
                itemsList.Clear();
                Console.WriteLine("\nArchivo Excel guardado exitosamente.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al guardar el archivo Excel: {ex.Message}");
            }
        }
    }
}