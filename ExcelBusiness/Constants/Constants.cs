namespace ExcelBusiness.Constants
{
    public class Constants
    {
       public static readonly string FilePathQuestionMessage = "Укажите путь до файла типа .xls или .xlsx с данными: ";
       public static readonly string FilePathErrorMessage = "Файл по указанному пути не найден!";
       public static readonly string FilePathSuccessMessage = "Файл успешно найден!";
       public static readonly string FileSheetsNamesMessage = "Файл содержит следующие таблицы:";
       public static readonly string FileSheetsNotFoundMessage = "Файл не содержит таблиц!";
       public static readonly string AppAvailableOptionsMessage = "Доступные возможности приложения:";
       public static readonly string OptionChooseHelperMessage = "Для выбора функции введите ее наименование: ";
       public static readonly string OptionChooseErrorMessage = "Неверное наименование функции!";
       public static readonly string ProductNameQuestionMessage = "Введите наименование товара: ";
       public static readonly string ProductNameErrorMessage = "Некорректное наименование товара!";
       public static readonly string ProductNotFoundMessage = "Указанный товар не найден!";
       public static readonly string OrderNotFoundMessage = "Заказов по данному товару не найдено!";
       public static readonly string DateRangeMessage = "Введите диапазон дат в формате dd.mm.yyyy-dd.mm.yyyy:";
       public static readonly string DateRangeErrorMessage = "Необходимо ввести диапазон дат в формате dd.mm.yyyy-dd.mm.yyyy!";
       public static readonly string DateRangeNotFoundMessage = "Заказы за указанный период не найдены!";
       public static readonly string AllClientsListMessage = "Таблица всех клиентов:";
       public static readonly string ClientsOrgToChangeMessage = "Введите название организации клиента для изменения: ";
       public static readonly string ClientsOrgErrorMessage = "Необходимо ввести корректное наименование организации!";
       public static readonly string ClientsOrgNewPersonMessage = "Введите новое наименование контактного лица организации: ";
       public static readonly string ClientsOrgNewPersonSuccessMessage = "Переименование контактного лица организации прошло успешно!";
       public static readonly List<string> appOptions = new List<string>() {"Вывод информации о клиентах по наименованию товара",
                                                                       "Запрос на изменение контактного лица клиента",
                                                                        "Запрос на определение золотого клиента"
       };
    }
}
