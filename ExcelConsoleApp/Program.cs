using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            String fileLocation = @"C:\Git\Excel97COM\ExcelConsoleApp\bin\Debug\Sample2.xls";
            ExcelReader _ExcelReader = new ExcelReader(fileLocation);        
            WorkSheetReader workSheetReader = new WorkSheetReader(_ExcelReader.getWorkSheet());
           
            Console.WriteLine(workSheetReader.ToString());

            foreach(ValueLocation vl in workSheetReader.mapValues())
            {
                String w_value = workSheetReader.getValue(vl.row, vl.col);
                if(w_value!= "")
                    Console.WriteLine("[" + vl.row + "," + vl.col + "] = " + w_value);
                     
            }
            workSheetReader.printValues();
          
            Console.ReadKey();
        }
    }
}
