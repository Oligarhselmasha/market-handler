using ClosedXML.Excel;

class Program
{
    

    static void Main(string[] args)
    {   
        Handler hendler = new ();
        hendler.Menu();
    }
}

class Handler
{
    private string? path;
    private readonly List<Product> products = new();
    private readonly List<Client> clients = new ();
    private readonly List<Request> requests = new ();

    /// <summary>
    /// Метод-обертка. Позволяет обработать действие, введенное пользователем в консоле.
    /// </summary>
    public void Menu()
    {
        string? action;
        while (true)
        {
            Console.WriteLine("Введите действие:");
            Console.WriteLine("1 - Загрузить файл");
            Console.WriteLine("2 - Получить список всех товаров");
            Console.WriteLine("3 - Получить список клиентов, заказавших товар");
            Console.WriteLine("4 - Получить список всех клиентов");
            Console.WriteLine("5 - Изменить контактное лицо клиента");
            Console.WriteLine("6 - Получить золотого клиента");
            Console.WriteLine("7 - Получить клиента с наибольшим количеством заказов");
            Console.WriteLine("0 - Выход");
            action = Console.ReadLine();

            switch (action)
            {
                case "1":
                    GetFile();
                    break;
                case "2":
                    GetProductsList();
                    break;
                case "3":
                    GetClientsByProductName();
                    break;
                case "4":
                    GetClientsList();
                    break;
                case "5":
                    ChangeContact();
                    break;
                case "6":
                    GetGoldClient();
                    break;
                case "7":
                    GetBestClientByDate();
                    break;
                case "0":
                    Console.WriteLine("Приятно было с Вами поработать, до свидания!");
                    return;
                default:
                    Console.WriteLine("Вы ввели неверную команду, попробуйте еще раз!");
                    Console.WriteLine();
                    break;
            }
        }
    }

    /// <summary>
    /// Метод-обертка для метода обработки файла. Проверяет есть файл по указанному пути в системе и, в случае обнаружения, запускает обработку.
    /// </summary>
    private void GetFile()
    {
        Console.WriteLine("Укажите путь до файла в формате '*путь_до_файла*Практическое задание для кандидата.xlsx'");
        path = Console.ReadLine(); 
        if(File.Exists(path)){
            MakeLists();
        }
        else
        {
            Console.WriteLine("Файл не существует!");
        }
    }

    /// <summary>
    /// Обработка файла
    /// </summary>
    private void MakeLists()
    {
        try
        {
            using (XLWorkbook wb = new (path))
            {
                ClearLists();
                if (!wb.Worksheets.TryGetWorksheet("Товары", out IXLWorksheet worksheetProd))
                {
                    Console.WriteLine("В указанном файле отсутствует лист 'Товары', попробуйте загрузить другой файл");
                    return;
                }

                IXLRows rows = worksheetProd.Rows();
                int rowsCount = rows.Count();
                for (int i = 2; i <= rowsCount; i++)
                {
                    Product product = new Product();
                    if (int.TryParse(worksheetProd.Cell(i, 1).Value.ToString(), out int Code))
                    {
                        product.Code = Code;
                    }
                    else
                    {
                        continue;
                    }
                    product.Name = worksheetProd.Cell(i, 2).Value.ToString();
                    product.UM = worksheetProd.Cell(i, 3).Value.ToString();
                    if (decimal.TryParse(worksheetProd.Cell(i, 4).Value.ToString(), out decimal Cost))
                    {
                        product.Cost = Cost;
                    }
                    if (product != null)
                    {
                        products.Add(product);

                    }
                }

                if (!wb.Worksheets.TryGetWorksheet("Клиенты", out IXLWorksheet worksheetCln))
                {
                    Console.WriteLine("В указанном файле отсутствует лист 'Клиенты', попробуйте загрузить другой файл");
                    return;
                }
                rows = worksheetCln.Rows();
                rowsCount = rows.Count();

                for (int i = 2; i <= rowsCount; i++)
                {
                    Client client = new();
                    if (int.TryParse(worksheetCln.Cell(i, 1).Value.ToString(), out int Code))
                    {
                        client.Code = Code;
                    }
                    else
                    {
                        continue;
                    }
                    client.Name = worksheetCln.Cell(i, 2).Value.ToString();
                    client.Adress = worksheetCln.Cell(i, 3).Value.ToString();
                    client.Contact = worksheetCln.Cell(i, 4).Value.ToString();
                    if (client != null)
                    {
                        clients.Add(client);

                    }
                }

                if (!wb.Worksheets.TryGetWorksheet("Заявки", out IXLWorksheet worksheetReq))
                {
                    Console.WriteLine("В указанном файле отсутствует лист 'Заявки', попробуйте загрузить другой файл");
                    return;
                }

                rows = worksheetCln.Rows();
                rowsCount = rows.Count();

                for (int i = 2; i <= rowsCount; i++)
                {
                    Request request = new();

                    if (int.TryParse(worksheetReq.Cell(i, 1).Value.ToString(), out int Code))
                    {
                        request.Code = Code;
                    }
                    else
                    {
                        continue;
                    }
                    if (int.TryParse(worksheetReq.Cell(i, 2).Value.ToString(), out int ProductCode))
                    {
                        request.ProductCode = ProductCode;
                    }
                    if (int.TryParse(worksheetReq.Cell(i, 3).Value.ToString(), out int ClientCode))
                    {
                        request.ClientCode = ClientCode;
                    }
                    if (int.TryParse(worksheetReq.Cell(i, 4).Value.ToString(), out int RequestNumber))
                    {
                        request.RequestNumber = RequestNumber;
                    }
                    if (int.TryParse(worksheetReq.Cell(i, 5).Value.ToString(), out int RequiredQty))
                    {
                        request.RequiredQty = RequiredQty;
                    }
                    if (DateTime.TryParseExact(worksheetReq.Cell(i, 6).Value.ToString(), "dd.MM.yyyy h:mm:ss",
                        System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime DeployDate))
                    {
                        request.DeployDate = DeployDate;
                    }
                    if (request != null)
                    {
                        requests.Add(request);
                    }
                }
            }
            Console.WriteLine();
            Console.WriteLine("Файл успешно считан!");
        }
        catch (Exception ex)
        {
            Console.WriteLine("При обработки файла {0} произошла ошибка! Удостоверьтесь в корректности файла.", path);
        }
    }

    /// <summary>
    /// Очистка переменных класса - сущностей полученных при считывании файла. Срабатывает при загрузке нового файла.
    /// </summary>
    private void ClearLists()
    {
        products.Clear();
        clients.Clear();
        requests.Clear();
    }

    /// <summary>
    /// Получение списка товаров с последующим выводом в консоль
    /// </summary>
    private void GetProductsList()
    {
        if (!IsValidRequest())
        {
            return;
        }  
           Console.WriteLine("Код_товара Наименование Ед.измерения Цена");
           products.ForEach(p => Console.WriteLine("{0} {1} {2} {3}", p.Code, p.Name, p.UM, p.Cost));
           Console.WriteLine();
    }

    /// <summary>
    /// Получение списка клиентов с последующим выводом в консоль
    /// </summary>
    private void GetClientsList()
    {
        if (!IsValidRequest())
        {
            return;
        }
            Console.WriteLine("Код_клиента Наименование Адрес Контактное_лицо");
            clients.ForEach(c => Console.WriteLine("{0} {1} {2} {3}", c.Code, c.Name, c.Adress, c.Contact));
            Console.WriteLine();
    }

    /// <summary>
    /// Получение получение информации по клиентам по полученному в методе наименованию продукта 
    /// </summary>
    private void GetClientsByProductName()
    {
        if (!IsValidRequest())
        {
            return;
        }
        string? productName;
        Console.WriteLine("Введите наименование товара, по которому Вы хотите получить информацию");
        productName = Console.ReadLine();
        
        Product? product = products.Where(p=>p.Name == productName).FirstOrDefault();
        if (product != null)
        {
            var orders = from r in requests
                              join p in products on r.ProductCode equals p.Code
                              join c in clients on r.ClientCode equals c.Code
                              where p.Name == productName
                              select new { ClientName = c.Name, Count = r.RequiredQty, Total = r.RequiredQty*p.Cost, r.DeployDate };
            if (!orders.Any())
            {
                Console.WriteLine("По товару {0} еще не было заявок, попробуйте выбрать другой!", productName);
                Console.WriteLine();
            }
            else
            {
                Console.WriteLine("Клиент Количество Сумма Дата");
                foreach (var order in orders)
                {
                    Console.WriteLine("{0} {1} {2} {3}", order.ClientName, order.Count, order.Total, order.DeployDate);
                }
                Console.WriteLine();
            }
        }
        else
        {
            Console.WriteLine("Информации по данному товару нет, со списком товаров можно ознакомиться во втором пункте меню");
        }
    }

    /// <summary>
    /// Изменение контактного лица у клиента по полученному в методе коду клиента
    /// </summary>
    private void ChangeContact()
    {
        if (!IsValidRequest())
        {
            return;
        }
        while (true)
        {
            Console.WriteLine("Введите код клиента, у которого Вы хотите сменить контактное лицо");
            if (!int.TryParse(Console.ReadLine(), out int clientCode))
            {
                Console.WriteLine("Введенный код клиента должен быть числом, попробуйте еще раз!");
                continue;
            }
            Client client = clients.Where(c => c.Code == clientCode).FirstOrDefault();
            if (client == null)
            {
                Console.WriteLine("Указанного клиента код {0} не существует!", clientCode);
                return;
            }
            else
            {
                Console.WriteLine("Введите новое контактное лицо для клиента {0}", clientCode);
                string newContactName = Console.ReadLine();
                try
                {
                    using (XLWorkbook wb = new(path))
                    {
                        IXLWorksheet sheet = wb.Worksheets.Worksheet("Клиенты");
                        int row = sheet.Rows().Count();
                        for (int i = 1; i <= row; i++)
                        {
                            if (sheet.Cell(i, 1).Value.ToString() == clientCode.ToString())
                            {
                                sheet.Cell(i, 4).Value = newContactName;
                            }
                        }
                        wb.Save();
                    }
                    MakeLists();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Возникли проблемы с доступом к исходному файлу");
                    Console.WriteLine();
                }
                
               return;
            }
        }
    }

    /// <summary>
    /// Получение золотого клиента, т.е. клиента совершившего завки на наибольшую сумму за все время
    /// </summary>
    private void GetGoldClient()
    {
        if (!IsValidRequest())
        {
            return;
        }

        var goldClient =
        requests
            .GroupBy(r => r.ClientCode)
            .Select(c => new
            {
                ClientId = c.Key,
                Total = c.Sum(r => r.RequiredQty * products.First(p => p.Code == r.ProductCode).Cost),
                Name = clients.First(n => n.Code == c.Key).Name

            })
            .OrderByDescending(x => x.Total)
            .First();
        Console.WriteLine("Золотой клиент - {0} с общей суммой заказов {1} рублей.", goldClient.Name, goldClient.Total);
        Console.WriteLine();
    }

    /// <summary>
    /// Получение получение лучшего клиента по дате. В методе две ветки - для получения по году и по месяцу
    /// </summary>
    private void GetBestClientByDate()
    {
        if (!IsValidRequest())
        {
            return;
        }
        string? cause;
        while (true)
        {
            Console.WriteLine("Укажите размер периода, за который Вы хотите вывести самого лучшего клиента");
            Console.WriteLine("Введите 'м' если за месяц");
            Console.WriteLine("Введите 'г' если за год");
            cause = Console.ReadLine();
            Console.WriteLine();
            switch (cause)
            {
                case "м":
                    Console.WriteLine("Введите год");
                    if (int.TryParse(Console.ReadLine(), out int year))
                    {
                        if (year > 0)
                        {
                            Console.WriteLine("Введите месяц (от 1 до 12, где 1 - это январь, а 12 - это декабрь)");
                            if (int.TryParse(Console.ReadLine(), out int month))
                            {
                                if (month >= 1 && month <= 12)
                                {
                                    try
                                    {
                                        var client = requests
                                                              .Where(r => r.DeployDate.Year == year && r.DeployDate.Month == month)
                                                              .GroupBy(r => r.ClientCode)
                                                              .Select(c => new
                                                                                {
                                                                                  ClientId = c.Key,
                                                                                  Name = clients.First(n => n.Code == c.Key).Name,
                                                                                  Count = c.Count()
                                                                                })
                                                              .OrderByDescending(x => x.Count)
                                                              .First();
                                        Console.WriteLine("Клиентом за указанный период времени за {0} год, {1} месяц явялется {2} (кол-во заказов" +
                                            " {3})", year, month, client.Name, client.Count);
                                        return;
                                    }
                                    catch
                                    {
                                        Console.WriteLine("По указанным данным невозможно определить клиента!");
                                        Console.WriteLine();
                                        return;
                                    }
                                }
                            }
                        }
                    }
                    Console.WriteLine("Некорректный период!");
                    Console.WriteLine();
                    continue;
                case "г":
                    Console.WriteLine("Введите год");
                    if (int.TryParse(Console.ReadLine(), out year))
                    {
                        if (year > 0)
                        {
                            try
                            {
                                var client = requests
                             .Where(r => r.DeployDate.Year == year)
                             .GroupBy(r => r.ClientCode)
                             .Select(c => new
                             {
                                 ClientId = c.Key,
                                 Name = clients.First(n => n.Code == c.Key).Name,
                                 Count = c.Count()

                             })
                             .OrderByDescending(x => x.Count)
                             .First();
                                Console.WriteLine("Клиентом за указанный период времени за {0} год, явялется {1} (кол-во заказов" +
                                                     " {2})", year, client.Name, client.Count);
                                return;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("По указанным данным невозможно определить клиента!");
                                Console.WriteLine();
                                return;
                            }

                        }
                    }
                    Console.WriteLine("Некорректный период!");
                    Console.WriteLine();
                    continue;
                default:
                    continue;
            }
        }    
    }

    /// <summary>
    /// Метод определеяет, что сущности, хранящиеся на уровне класса не пустые, т.е. файл загружен
    /// </summary>
    private bool IsValidRequest()
    {
        if (products.Count == 0 || requests.Count == 0 || clients.Count == 0)
        {
            Console.WriteLine("Списки не заполнены. Возможно Вы не загрузили файл");
            Console.WriteLine("Выполните первый пункт в меню");
            Console.WriteLine();
            return false;
        }
        else return true;
    }
}

/// <summary>
/// Класс соответствующей сущности "Товар"
/// </summary>
class Product
{
    public int Code { get; set; }
    public string? Name { get; set; }
    public string? UM { get; set; }
    public decimal Cost { get; set; }
}

/// <summary>
/// Класс соответствующей сущности "Клиент"
/// </summary>
class Client
{
    public int Code { get; set; }
    public string? Name { get; set; }
    public string? Adress { get; set; }
    public string? Contact { get; set; }
}

/// <summary>
/// Класс соответствующей сущности "Заявка"
/// </summary>
class Request
{
    public int Code { get; set; }
    public int ProductCode { get; set; }
    public int ClientCode { get; set; }
    public int RequestNumber { get; set; }
    public int RequiredQty { get; set; }
    public DateTime DeployDate { get; set; }
}