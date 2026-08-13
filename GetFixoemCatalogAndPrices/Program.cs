using GetFixoemCatalogAndPrices.Dto;
using GetFixoemCatalogAndPrices.Service;
using GetFixoemCatalogAndPrices.Utilities;
using System;
using System.Text.Json;


namespace GetFixoemCatalogAndPrices
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string apiKey = "7x7X2W9L3A";
            int maxResults;

            Console.WriteLine("Introduce el nombre del articulo que deseas buscar");
            string searchParameter = Console.ReadLine();

            var client = new ApiClientService();
            
            string apiResponse = client.GetCatalogBySearchParameter(apiKey, Common.HtmlEncodeStringValue(searchParameter)).GetAwaiter().GetResult();
            JsonDocument document = JsonDocument.Parse(apiResponse);
            maxResults = document.RootElement.GetProperty("totalItems").GetInt32();

            string result = client.GetCatalogBySearchParameter(apiKey, Common.HtmlEncodeStringValue(searchParameter), maxResults).GetAwaiter().GetResult();
            var productResponse = Common.DeserializeResponse<ProductResponse>(result);

            string filePath = @"C:\Users\quetz\OneDrive\Escritorio\" + searchParameter + ".xlsx";
            Excel.CreateFile(filePath);

            var items = Adapter.ConvertProductResponseItemsToItemList(productResponse);
            Excel.PopulateExcelFile(filePath, items);

            Console.WriteLine("Archivo Excel creado exitosamente en: " + filePath);
        }
    }
}
