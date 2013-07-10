Read and Write Excel File Dynamically
=======
In this tip, I decided to describe two parts, one part related to how I can read Excel file and the second part describes how I can generate a class dynamically and fill it.

When an Excel file is passed to the code, the first row of Excel file chooses as a class properties that we want to generate it. This functionality is implemented in GetRowValues static method in Wisgance.Office.Excel.Reader.Read class:

    private static List<string> GetColumnValues(Stream file, string sheetName, string reference)
            {
                var result = new List<string>();
     
                //Read Excel File by OpenXml Library
                using (var document = SpreadsheetDocument.Open(file, false))
                {
                    var workbook = document.WorkbookPart;
     
                    var theSheet = workbook.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName) ??
                                     workbook.Workbook.Descendants<Sheet>().FirstOrDefault(sheet => true);
     
                    var worksheet = (WorksheetPart)(workbook.GetPartById(theSheet.Id));
     
                    var cells = worksheet.Worksheet.Descendants<Cell>().Where(c => GetCellCol(c.CellReference).ToUpper() == reference);
     
                    //get cell data by calling ExtractCellValue function
                    result.AddRange(from theCell in cells where theCell != null select ExtractCellValue(theCell, workbook));
                }
     
                return result;
            } 

In this function and some others like GetRowValues or GetCellData exists argument named references. This argument gets Excel cell address, for example for first row data, we send "1" but for first column, we send "A".

After extracting first row data, we passed these data to the CreateNewObject function in Wisgance.Reflection.WisganceTypeGenerator class.

     public static object CreateNewObject(List<FieldMask> props)
            {
                var myType = CompileResultType(props);
                var myObject = Activator.CreateInstance(myType);
                return myObject;
            } 

In the first line, we generate a type via extracted data from Excel (CompileResultType) . In the second line, we create an instance of this class and return it.

    while (true)
                {
                    var obj = WisganceTypeGenerator.CreateNewObject(objProp);
                    var values = GetRowValues(stream, "", (k + 1).ToString());
                    if (!values.Any())
                        break;
                    for (var c = 0; c < objProp.Count; c++)
                    {
                        try
                        {
                            var propertyInfo = obj.GetType().GetProperty(objProp[c].FieldName);
                            propertyInfo.SetValue(obj, values[c], null);
                        }
                        catch (Exception) { }
                    }
                    result.Add(obj);
                    k++;
                } 

While true loop, this part is for reading data till it reaches the empty row. In each iteration, an object instantiates and fills properties value by selected row cell values.

For writing a list of objects in Excel, call Do function in Wisgance.Office.Excel.Writer.Write.

    if (headerNames == null)
    {
      headerNames = new ExcelHeaderList();
    	foreach (var o in ObjUtility.GetPropertyInfo(objects[0]))
        	{
           		headerNames.Add(o, o);
          	}
    } 

If headerNames is null, automatically set the property name as an Excel header, and also you can pass headerList if you want to customize the header name.

You can download it in [Nuget][1] here


  [1]: https://nuget.org/packages/ExcelDynamicReaderWriter/
