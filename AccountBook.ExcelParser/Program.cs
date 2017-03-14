using System;
using System.Collections.Generic;

namespace AccountBook.ExcelParser
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var accountBookEntries = new List<AccountBookDto>();

            var excelParser = new ExcelParser();
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2013.xlsx", @"Alexander"));
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2013.xlsx", @"Verena"));
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2014.xlsx", @"Alexander"));
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2014.xlsx", @"Verena"));
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2015.xlsx", @"Alexander"));
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2015.xlsx", @"Verena"));
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2016.xlsx", @"Alexander"));
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2016.xlsx", @"Verena"));
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2017.xlsx", @"Alexander"));
            accountBookEntries.AddRange(excelParser.Parse(@"D:\Temp\Haushaltsbuch 2017.xlsx", @"Verena"));

            var entryUploader = new EntryUploader();
            entryUploader.UploadEntries(accountBookEntries, @"http://localhost:8081/api/update");

            Console.WriteLine("Finished! Processed {0} entries. Press any key to exit.", accountBookEntries.Count);
            Console.ReadKey();
        }
    }
}