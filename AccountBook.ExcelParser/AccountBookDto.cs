using Newtonsoft.Json;
using System;

namespace AccountBook.ExcelParser
{
    public class AccountBookDto
    {
        [JsonProperty(PropertyName = "bookingDate")]
        public DateTime BookingDate { get; set; }

        [JsonProperty(PropertyName = "user")]
        public string User { get; set; }

        [JsonProperty(PropertyName = "title")]
        public string Title { get; set; }

        [JsonProperty(PropertyName = "amount")]
        public double Amount { get; set; }

        [JsonProperty(PropertyName = "incomeExpense")]
        public string IncomeExpense { get; set; }

        [JsonProperty(PropertyName = "recordType")]
        public string RecordType { get; set; }

        public override string ToString()
        {
            return string.Format("User: {0}{1}Booking Date: {2}{3}Title: {4}{5}Amount: {6}{7}IncomeExpense: {8}{9}Record Type: {10}{11}", 
                User, Environment.NewLine, 
                BookingDate, Environment.NewLine, 
                Title, Environment.NewLine, 
                Amount, Environment.NewLine, 
                IncomeExpense, Environment.NewLine, 
                RecordType, Environment.NewLine);
        }
    }
}