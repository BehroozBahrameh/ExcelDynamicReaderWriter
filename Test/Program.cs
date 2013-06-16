using System;
using System.IO;
using Wisgance.Office.Excel.Reader;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var starTime = DateTime.Now;
            Console.WriteLine("Start Reading Data");
            var stream = new MemoryStream(File.ReadAllBytes(@"G:\20q\1.xlsx"));
            var data = Read.ReadObjFromExel(stream);
            Console.WriteLine("{0} : {1}", "Data Reading Finishd", DateTime.Now - starTime);
            Console.WriteLine("Start Writing Data");
            var result = new Wisgance.Office.Excel.Writer.Write().Do(data, "BehroozSheet", null);
            result.Seek(0, SeekOrigin.Begin);
            var ms = (MemoryStream)result;
            var file = new FileStream(@"G:\WisganceResult.xlsx", FileMode.Create, System.IO.FileAccess.Write);
            var bytes = new byte[ms.Length];
            ms.Read(bytes, 0, (int)ms.Length);
            file.Write(bytes, 0, bytes.Length);
            file.Close();
            ms.Close();
            Console.WriteLine("Start Writing Finished");
            Console.ReadKey();
        }
    }
}
