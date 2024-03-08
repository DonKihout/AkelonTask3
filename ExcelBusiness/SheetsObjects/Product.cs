using Ganss.Excel;

namespace ExcelBusiness.SheetsObjects
{
    public class Product
    {
        [Column("Код товара")]
        public int ProductCode { get; set; }

        [Column("Наименование")]
        public string ProductName { get; set; }

        [Column("Ед. измерения")]
        public string MeasuretUnit { get; set; }

        [Column("Цена товара за единицу")]
        public decimal Cost { get; set; }
    }
}
