using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace AccountBook.ExcelParser
{
    public class ExcelParser
    {
        private const int StartColumnAlexander = 0;
        private const int StartColumnVerena = 8;
        private const int StartRow = 2;

        public IList<AccountBookDto> Parse(string fileName, string user)
        {
            XSSFWorkbook wb;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(file);
            }

            var accountBookEntries = new List<AccountBookDto>();

            for (int i = 1; i < wb.NumberOfSheets; i++)
            {
                var sheet = wb.GetSheetAt(i);
                if (sheet != null)
                {
                    DateTime lastValidBookingDate = new DateTime();
                    for (int row = StartRow; row <= sheet.LastRowNum; row++)
                    {                       
                        var currentRow = sheet.GetRow(row);
                        if (currentRow == null)
                        {
                            continue;
                        }

                        var cell = user == "Alexander" ? StartColumnAlexander : StartColumnVerena;
                        double? amount = null;
                        string incomeExpense = string.Empty;

                        // Booking Date
                        var bookingDate = currentRow.GetCell(cell, MissingCellPolicy.CREATE_NULL_AS_BLANK).DateCellValue;
                        if (bookingDate == null || bookingDate.Year == 0001)
                        {
                            bookingDate = lastValidBookingDate;
                        }
                        else
                        {
                            lastValidBookingDate = bookingDate;
                        }

                        // Title
                        cell++;
                        var title = currentRow.GetCell(cell, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue;
                        if (string.IsNullOrEmpty(title))
                        {
                            // Ignore rows with empty title
                            continue;
                        }

                        // Expense
                        cell++;

                        var expense = currentRow.GetCell(cell, MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue;
                        if (expense != 0)
                        {
                            amount = Math.Round(expense, 2);
                            incomeExpense = "E";
                        }

                        // Income
                        cell++;
                        var income = currentRow.GetCell(cell, MissingCellPolicy.CREATE_NULL_AS_BLANK).NumericCellValue;
                        if (income != 0)
                        {
                            amount = Math.Round(income, 2); ;
                            incomeExpense = "I";
                        }

                        // Voucher Expense
                        cell++;

                        // Voucher Income
                        cell++;

                        // Record Type
                        cell++;
                        var recordType = currentRow.GetCell(cell, MissingCellPolicy.CREATE_NULL_AS_BLANK).StringCellValue;
                        if (recordType == "*" || recordType == "X")
                        {
                            // Ignore old voucher and special investment incomes and expenses
                            continue;
                        }
                        if (bookingDate.Date >= new DateTime(2013, 01, 01).Date && bookingDate.Date <= new DateTime(2015, 05, 31).Date)
                        {
                            // Map record types A --> M and M --> F for entries starting from 2013 and ending in May 2014
                            var newRecordType = recordType;
                            if (recordType == "A")
                            {
                                newRecordType = "M";
                            }
                            if (recordType == "M")
                            {
                                newRecordType = "F";
                            }
                            recordType = newRecordType;
                        }

                        var accountBookDto = new AccountBookDto()
                        {
                            User = user,
                            BookingDate = bookingDate,
                            Title = title,
                            Amount = amount.GetValueOrDefault(),
                            IncomeExpense = incomeExpense,
                            RecordType = recordType
                        };

                        accountBookEntries.Add(accountBookDto);

                        Console.WriteLine(accountBookDto);
                    }
                }
            }

            return accountBookEntries;
        }
    }
}