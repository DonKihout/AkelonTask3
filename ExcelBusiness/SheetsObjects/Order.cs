using Ganss.Excel;

namespace ExcelBusiness.SheetsObjects
{
    public class Order
    {
        [Column("Код заявки")]
        public int OrderCode { get; set; }

        [Column("Код товара")]
        public int ProductCode { get; set; }

        [Column("Код клиента")]
        public int ClientCode { get; set; }

        [Column("Номер заявки")]
        public int OrderNumber { get; set; }

        [Column("Требуемое количество")]
        public int RequestedCount { get; set; }

        [Column("Дата размещения")]
        public DateOnly OrderDate { get; set; }
    }
}
