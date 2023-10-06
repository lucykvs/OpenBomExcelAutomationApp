using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json.Serialization;
using Newtonsoft.Json;
using System.Numerics;
using System.Security.Policy;
using Microsoft.VisualBasic;
using System.Net.Http.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace OpenBomWpf
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static HttpClient client = new HttpClient();

        public MainWindow()
        {
            InitializeComponent();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            string accessToken = await GetAccessTokenAsync();
            string catalogJson = await GetCatalogAsync(accessToken);
            List<CatalogItem> catalogItems = new();

            if (catalogJson != "")
            {
                catalogItems = ConvertToCatalogItems(catalogJson);
            }

            PopulateExcelSheet(catalogItems);
        }

        static void PopulateExcelSheet(List<CatalogItem> catalogItems)
        {
            var application = new Excel.Application();
            var workbook = application.Workbooks.Add();
            var sheet = (Excel.Worksheet)workbook.Sheets.Add();

            Excel.Range catalogItemRange = GetCatalogItemRange(sheet, catalogItems);
            object[,] catalogItemMatrix = GetCatalogItemMatrix(catalogItems);

            catalogItemRange.Value2 = catalogItemMatrix;

            sheet.SaveAs("TODO: replace");

            // for memory cleanup/garbage collection
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            if (workbook != null)
            {
                workbook.Close();
                workbook = null;
            }
            if (application != null)
            {
                application.Quit();
                application = null;
            }
        }

        static object[,] GetCatalogItemMatrix(List<CatalogItem> catalogItems)
        {
            int rows = catalogItems.Count;
            int columns = 6;
            object[,] catalogItemValues = new object[rows, columns];

            for (int i = 0; i < rows; i++)
            {
                catalogItemValues[i, 0] = catalogItems[i].PartNumber;
                catalogItemValues[i, 1] = catalogItems[i].SupplierPartNumber;
                catalogItemValues[i, 2] = catalogItems[i].Vendor;
                catalogItemValues[i, 3] = catalogItems[i].Description;
                catalogItemValues[i, 4] = catalogItems[i].SizeOfKanban;
                catalogItemValues[i, 5] = catalogItems[i].URL;
            }

            return catalogItemValues;
        }

        static Excel.Range GetCatalogItemRange(Excel.Worksheet sheet, List<CatalogItem> catalogItems)
        {
            Cell rangeTopCell = new Cell(1, 1);
            Cell rangeBottomCell = new Cell(catalogItems.Count, 6);
            Range range = new Range(rangeTopCell, rangeBottomCell);

            return range.GetExcelRange(sheet);
        }

        static List<CatalogItem> ConvertToCatalogItems(string catalogJson)
        {
            var deserialized = JsonConvert.DeserializeObject<Catalog>(catalogJson);

            List<string> columns = deserialized.Columns;

            int partNumIndex = deserialized.Columns.IndexOf("Part Number");
            int supplierNumIndex = deserialized.Columns.IndexOf("Supplier Part Number");
            int VendorIndex = deserialized.Columns.IndexOf("Vendor");
            int descriptionIndex = 3;
            int sizeOfKanbanIndex = deserialized.Columns.IndexOf("Size of Kanban");
            int urlIndex = deserialized.Columns.IndexOf("URL");

            List <CatalogItem> catalogItems = new();

            foreach (var item in deserialized.Cells)
            {
                catalogItems.Add(new CatalogItem(item[partNumIndex], item[supplierNumIndex], item[VendorIndex], 
                    item[descriptionIndex], item[sizeOfKanbanIndex], item[urlIndex]));
            }

            return catalogItems;
        }

        static async Task<string> GetAccessTokenAsync()
        {
            string requestUrl = "https://developer-api.openbom.com/login";
            string apiKey = "TODO";
            string accessToken = "";

            // TODO: replace stringcontent

            var request = new HttpRequestMessage(HttpMethod.Post, requestUrl);
            request.Headers.Add("X-OpenBOM-AppKey", apiKey);
            request.Content = content;
            var response = await client.SendAsync(request);
            response.EnsureSuccessStatusCode();
            Console.WriteLine(await response.Content.ReadAsStringAsync());

            var responseContents = await response.Content.ReadAsStringAsync();
            var deserialized = JsonConvert.DeserializeObject<AccessTokenResponse>(responseContents);

            return deserialized.AccessToken;
        }


        static async Task<string> GetCatalogAsync(string accessToken)
        {
            string basePath = "https://developer-api.openbom.com/catalog/";
            string apiKey = "TODO";
            string catalogId = "TODO";

            string requestUrl = basePath + catalogId;

            client.DefaultRequestHeaders.Add("X-OpenBOM-AppKey", apiKey);
            client.DefaultRequestHeaders.Add("X-OpenBOM-AccessToken", accessToken);

            HttpResponseMessage response = await client.GetAsync(requestUrl);

            var b = 1;

            if (response.IsSuccessStatusCode)
            {
                var contents = await response.Content.ReadAsStringAsync();
                return contents;
            } else
            {
                return "";
            }
        }
    }

    class Catalog
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ModifiedDate { get; set; }
        public string CreatedDate { get; set; }
        public string ModifiedBy { get; set; }
        public string CreatedBy { get; set; }
        public List<string> Columns { get; set; }
        public List<List<string>> Cells { get; set; }
    }

    class CatalogItem
    {
        public string PartNumber { get; set; }
        public string SupplierPartNumber { get; set; }
        public string Vendor { get; set; }
        public string Description { get; set; }
        public string SizeOfKanban { get; set; }
        public string URL { get; set; }

        public CatalogItem(string partNumber, string supplierPartNumber, string vendor, string description, string sizeOfKanban, string url)
        {
            this.PartNumber = partNumber;
            this.SupplierPartNumber = supplierPartNumber;
            this.Vendor = vendor;
            this.Description = description;
            this.SizeOfKanban = sizeOfKanban;
            this.URL = url;
        }
    }

    class AccessTokenResponse
    {
        [JsonProperty("access_token")]
        public string AccessToken { get; set; }
    }
}