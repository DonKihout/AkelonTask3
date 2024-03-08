using Ganss.Excel;

namespace ExcelBusiness.SheetsObjects
{
    public class Client
    {
        [Column("Код клиента")]
        public int ClientCode { get; set; }

        [Column("Наименование организации")]
        public string OrganizationName { get; set; }

        [Column("Адрес")]
        public string Address { get; set; }

        [Column("Контактное лицо (ФИО)")]
        public string Person { get; set; }
    }
}
