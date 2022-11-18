namespace ConsoleApp2
{
    public class ExcelProductsModel
    {
        public string SearchString { get; set; }
        public List<ProductModel> Products { get; set; }
    }

    public class RespnoseModel
    {
        public DataModel Data { get; set; }
    }

    public class DataModel
    {
        public List<ProductModel> Products { get; set; }
    }

    public class ProductModel
    {
        public long Id { get; set; }
        public string Name { get; set; }
        public string Brand { get; set; }
        public long PriceU { get; set; }
        public int Feedbacks { get; set; }
    }
}
