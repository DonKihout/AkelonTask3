using ExcelBusiness.Constants;
using ExcelBusiness;

namespace AkelonTask3
{
    public class Program
    {
        
        static void Main(string[] args)
        {
            Console.WriteLine("Задание 3. Разработка решения с использованием документа Microsoft Excel");
            OpenFile();
        }

        public static void OpenFile() 
        {
            Console.Write(Constants.FilePathQuestionMessage);
            var filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                Console.WriteLine(Constants.FilePathErrorMessage);
                OpenFile();
            }
            else 
            {
                try 
                {
                    Console.WriteLine(Constants.FilePathSuccessMessage);
                    ExcelFile excelFile = new ExcelFile(filePath);
                    excelFile.OpenExcelFile();
                } 
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    OpenFile();
                }                
            }          
        }
    }
}