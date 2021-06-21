using GemBox.Spreadsheet;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;

namespace Task2
{
    class Program
    {
        static void Main(string[] args)
        {
            JsonToExcel jsonToExcel = new JsonToExcel();
            JsonData items = jsonToExcel.LoadJson();
            string path = jsonToExcel.CreateExcel(items);
            // saving file in project location in bin/debug/net folder
            Console.WriteLine("File Created: {0}", path);
        }

    }
}
