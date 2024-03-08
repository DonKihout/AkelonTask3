namespace ExcelBusiness.ResultObjects
{
    public class ClientsInfo
    {
        public int? ProductCount { get; set; }
        public decimal? SingleProductPrice { get; set; }
        public DateOnly? OrderDate { get; set; }
        public string? Person { get; set; }
        public string? Address { get; set; }
        public string? ClientName { get; set; }
        public string? ProductName { get; set; }
    }
}
