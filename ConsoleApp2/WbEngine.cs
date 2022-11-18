using Aspose.Cells;
using Newtonsoft.Json;

namespace ConsoleApp2
{
    public class WbEngine
    {
        public async Task<List<ProductModel>> GetDataByQueryTerm(string searchString)
        {
            var pageQueryParam = 1;
            var queryString2 = $"https://search.wb.ru/exactmatch/ru/common/v4/search?appType=1&couponsGeo=12,3,18,15,21&page={pageQueryParam}&curr=rub&dest=-1029256,-102269,-2162196,-1257786&emp=0&lang=ru&locale=ru&pricemarginCoeff=1.0&query={searchString}&resultset=catalog&sort=popular&spp=0&suppressSpellcheck=false";
            HttpClient client = new();
            HttpResponseMessage response = await client.GetAsync(queryString2);
            response.EnsureSuccessStatusCode();
            string responseBody = await response.Content.ReadAsStringAsync();

            var products = JsonConvert.DeserializeObject<RespnoseModel>(responseBody).Data.Products;

            return products;
        }

        public void WriteIntoExcel(List<ExcelProductsModel> excelProducts)
        {
            Workbook wkb = new();
            wkb.Worksheets.Clear();

            for (int i = 0; i < excelProducts.Count; i++)
            {
                var excelProduct = excelProducts[i];
                var sht = wkb.Worksheets.Add(excelProduct.SearchString);

                Cell a1 = sht.Cells["A1"];
                a1.PutValue("Title");

                Cell b1 = sht.Cells["B1"];
                b1.PutValue("Brand");

                Cell c1 = sht.Cells["C1"];
                c1.PutValue("Id");

                Cell d1 = sht.Cells["D1"];
                d1.PutValue("Feedback");

                Cell e1 = sht.Cells["E1"];
                e1.PutValue("Price");

                for (int j = 0; j < excelProduct.Products.Count; j++)
                {
                    var rowNumber = j + 1;
                    var product = excelProduct.Products[j];
                    Cell cell1 = sht.Cells[rowNumber, 0];
                    cell1.PutValue(product.Name);

                    Cell cell2 = sht.Cells[rowNumber, 1];
                    cell2.PutValue(product.Brand);

                    Cell cell3 = sht.Cells[rowNumber, 2];
                    cell3.PutValue(product.Id);

                    Cell cell4 = sht.Cells[rowNumber, 3];
                    cell4.PutValue(product.Feedbacks);

                    Cell cell5 = sht.Cells[rowNumber, 4];
                    cell5.PutValue(product.PriceU);
                }

                sht.AutoFitColumns();
            }
            
            wkb.Save("products.xlsx");
        }
    }
}
