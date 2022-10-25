using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordDocDataFill;

namespace TestDocFillConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {

            Dictionary<string, string> dic = new Dictionary<string, string>();

            dic.Add("[LastName]", "Bagaev");
            dic.Add("[FirstName]", "Vitaly");
            dic.Add("[PasspoertSN]", "1234 123456");

            try
            {
                IWordDocDataFill docDataFill = new WordDocDataFillInstance("C:\\Example\\ExampleTemplate.docx", "C:\\Example\\Result");

                docDataFill.FillDocument(dic, "TestFile1.docx");
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
