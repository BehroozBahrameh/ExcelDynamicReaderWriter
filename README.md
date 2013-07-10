#Read and Write Excel File Dynamically
##About
This project provide to you extracting data from excel documents as a list of object that dynamically.

##How to use
If want get list of object that dynamically created by excel header names, just call below function and send your file address!
    
    var stream = new MemoryStream(File.ReadAllBytes("YOUR_FILE_ADDRESS.xlsx"));

and also you can fill list of your object by sending list of object property name and excel related header in below function :
    
     ReadObjFromExel<T>(Stream stream, ExcelHeaderList pattern, string sheetName)
 
 - **stream** excel file
 - **pattern** list of related object property and excel columns
 - **sheetname** selected sheet of excel file for extracting data

pattern is a list of below class

    public class ExcelHeader
    {
        public string Key { get; set; }
        public string Value { get; set; }
        public ExcelHeaderType HeaderType { get; set; }
    }

alse you can get a list of object as excel file (function return type is Stream ), jest send your list to below function
    
    Wisgance.Office.Excel.Writer.Write().Do(yourListOfData, "SheetName", null);

First argement is your list, second one is sheetname that you want set in generated excel file and third one is culomn name.
if sent null for third argument, function put property name as column name.

-------------------
You can add this library by [nuget][1] in your project

  [1]: https://nuget.org/packages/ExcelDynamicReaderWriter/
