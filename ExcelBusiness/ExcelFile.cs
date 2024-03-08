using ExcelBusiness.ResultObjects;
using ExcelBusiness.SheetsObjects;
using Ganss.Excel;
using System.Data;
using Utulities;

namespace ExcelBusiness
{
    public class ExcelFile
    {
        private string filePath;
        private ExcelMapper excelMapper;
        private List<Product> products;
        private List<Order> orders;
        private List<Client> clients;
        private TablePrinter clientsInfoTableByProduct = new TablePrinter("Наименование организации", "Адрес", "Контактное лицо (ФИО)", 
                                                    "Продукт", "Цена товара за единицу", "Требуемое количество", "Дата размещения");
        private TablePrinter clientsTable = new TablePrinter("Код клиента", "Наименование организации", "Адрес",
                                                    "Контактное лицо (ФИО)");
        public ExcelFile(string filePath)
        {
            this.filePath = filePath;         
        }

        //Метод открытия и парсинга excel файла
        public void OpenExcelFile()
        {
            try
            {
                using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    excelMapper = new ExcelMapper(fileStream);
                }
                products = excelMapper.Fetch<Product>(sheetName: "Товары").ToList();
                orders = excelMapper.Fetch<Order>(sheetName: "Заявки").ToList();
                clients = excelMapper.Fetch<Client>(sheetName: "Клиенты").ToList();
                GetAppFunctions();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
        }

        //Метод вывода функций приложения
        public void GetAppFunctions()
        {
            Console.WriteLine(Constants.Constants.AppAvailableOptionsMessage);
            foreach (var option in Constants.Constants.appOptions)
            {
                Console.WriteLine(" - " + option);
            }
            Console.Write(Constants.Constants.OptionChooseHelperMessage);
            var optionName = Console.ReadLine();
            ChoiseHandler(optionName);
        }

        #region Метод получения информации о клиентах по продукту
        public void GetClientInfoByProduct() 
        {
            Console.Write(Constants.Constants.ProductNameQuestionMessage);
            var productName = Console.ReadLine();
            if (productName is not null && productName.Length > 0)
            {                
                var productCode = products.First(p => p.ProductName == productName).ProductCode.Equals(null) ? 
                                                            throw new Exception(Constants.Constants.ProductNotFoundMessage) 
                                                            :
                                                            products.First(p => p.ProductName == productName).ProductCode;
                var productOrders = orders.FindAll(o => o.ProductCode == productCode).Any() ? orders.FindAll(o => o.ProductCode == productCode)
                                                                                            : throw new Exception(Constants.Constants.OrderNotFoundMessage);
               
                List<ClientsInfo> clientsInfo = new List<ClientsInfo>();

                foreach (var order in productOrders)
                {
                    ClientsInfo client = new ClientsInfo()
                    {
                        ProductCount = order.RequestedCount,
                        SingleProductPrice = products.First(p => p.ProductCode == productCode).Cost,
                        OrderDate = order.OrderDate,
                        Person = clients.First(c => c.ClientCode == order.ClientCode).Person,
                        Address = clients.First(c => c.ClientCode == order.ClientCode).Address,
                        ClientName = clients.First(c => c.ClientCode == order.ClientCode).OrganizationName,
                        ProductName = productName
                    };
                    clientsInfo.Add(client);
                }

                foreach (var client in clientsInfo)
                {
                    clientsInfoTableByProduct
                                .AddRow(client.ClientName, client.Address, client.Person, client.ProductName, client.SingleProductPrice, client.ProductCount, client.OrderDate);
                }
                clientsInfoTableByProduct.Print();
                GetAppFunctions();
            }
            else 
            {
                Console.WriteLine(Constants.Constants.ProductNameErrorMessage);
                GetClientInfoByProduct();
            }
        }
        #endregion

        #region Метод изменения контактного лица организации
        public void ChangeContactPerson()
        {
            Console.WriteLine(Constants.Constants.AllClientsListMessage);
            foreach (var client in clients)
            {
                clientsTable.AddRow(client.ClientCode, client.OrganizationName, client.Address, client.Person);
            }
            clientsTable.Print();
            Console.Write(Constants.Constants.ClientsOrgToChangeMessage);

            try
            {
                var selectedOrg = Console.ReadLine();
                var client = clients.First(c => c.OrganizationName == selectedOrg) != null ? clients.First(c => c.OrganizationName == selectedOrg)
                                                                                            :
                                                                                            throw new Exception();
                clientsTable.AddRow(client.ClientCode, client.OrganizationName, client.Address, client.Person);
                clientsTable.Print();
                Console.WriteLine(Constants.Constants.ClientsOrgNewPersonMessage);
                var newOrgPersonName = Console.ReadLine();
                if (newOrgPersonName.Length == 0) throw new Exception();
                client.Person = newOrgPersonName;
                excelMapper.Save(filePath, clients, "Клиенты");
                Console.WriteLine(Constants.Constants.ClientsOrgNewPersonSuccessMessage);
                GetAppFunctions();
            }
            catch
            {
                Console.WriteLine(Constants.Constants.ClientsOrgErrorMessage);
                ChangeContactPerson();
            }

        }
        #endregion

        #region Метод получения "золотого" клиента
        public void GetGoldClientInfo()
        {
            Console.Write(Constants.Constants.DateRangeMessage);
            var dateRangeListString = Console.ReadLine().Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (dateRangeListString.Length != 2)
            {
                Console.WriteLine(Constants.Constants.DateRangeErrorMessage);
                GetGoldClientInfo();
            }
            try
            {
                var dateRangeList = new List<DateOnly>();
                foreach (var date in dateRangeListString)
                {
                    dateRangeList.Add(DateOnly.ParseExact(date, "dd.MM.yyyy",
                                       System.Globalization.CultureInfo.InvariantCulture));
                }
                dateRangeList.Sort();
                var ordersInRange = orders.Where(o => o.OrderDate > dateRangeList[0] && o.OrderDate < dateRangeList[1]);
                if (!ordersInRange.Any()) 
                {
                    Console.WriteLine(Constants.Constants.DateRangeNotFoundMessage);
                    GetGoldClientInfo();
                } 
                var goldClientCode = ordersInRange.GroupBy(o => o.ClientCode).OrderByDescending(grp => grp.Count())
                                                                                    .Select(grp => grp.Key).First();
                var goldClient = clients.First(c => c.ClientCode == goldClientCode);
                clientsTable.AddRow(goldClient.ClientCode, goldClient.OrganizationName, goldClient.Address, goldClient.Person);
                clientsTable.Print();
                GetAppFunctions();
            }
            catch
            {
                Console.WriteLine(Constants.Constants.DateRangeErrorMessage);
                GetGoldClientInfo();
            }
            
        }
        #endregion

        //Метод обработки пользовательского ввода по функциям приложения
        public void ChoiseHandler(string userChoise)
        {
            try {
                switch (userChoise)
                {
                    case "Вывод информации о клиентах по наименованию товара":
                        GetClientInfoByProduct();
                        break;
                    case "Запрос на изменение контактного лица клиента":
                        ChangeContactPerson();
                        break;
                    case "Запрос на определение золотого клиента":
                        GetGoldClientInfo();
                        break;
                    default:
                        Console.WriteLine(Constants.Constants.OptionChooseErrorMessage);
                        GetAppFunctions();
                        break;
                }
            }
            catch (Exception ex) { 
                Console.WriteLine(ex.Message);
                GetAppFunctions();
            }        
        }     
    }
}
