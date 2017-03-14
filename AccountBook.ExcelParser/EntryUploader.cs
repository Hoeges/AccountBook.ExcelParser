using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace AccountBook.ExcelParser
{
    public class EntryUploader
    {
        public void UploadEntries(IList<AccountBookDto> entries, string url)
        {
            foreach (var entry in entries)
            {
                Upload(entry, url);
            }
        }

        private void Upload(AccountBookDto entry, string url)
        {
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";

            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                var entryAsJson = JsonConvert.SerializeObject(entry);
                streamWriter.Write(entryAsJson);
            }

            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();
                Console.WriteLine(result);
            }
        }
    }
}