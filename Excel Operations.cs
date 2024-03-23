using System.Reflection;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace ExcelOperations
{
	/*
	This class allows you to: 
	Export a database query result into an excel file 
	and
	Import data of excel file into database table 
	in a completely generic way.
	
	Import: First line of excel file must contain column names of database table and backend data model. 
	Orders of columns in excel file is not important, also missing columns will be tried to inserted as null if database allows.
	First empty cell in the first row will be accepted as last column.
	
	Export: Just call the function with any data model.
	*/
    public static class ExcelExportAndImport
    {
        public static byte[] ExportDataToExcelFile<T>(string sheetName, IEnumerable<T> data)
        {
            using var workBook = new XLWorkbook();

            var workSheet = workBook.Worksheets.Add(sheetName);

            Type objectType = typeof(T);
            PropertyInfo[] properties = objectType.GetProperties();

			int k = 1;
            foreach (PropertyInfo property in properties)
            {
                workSheet.Cell(1, k).Value = Regex.Replace(property.Name, "(\\B[A-Z])", " $1");
                workSheet.Cell(1, k).Style.Font.FontSize = 12;
                workSheet.Cell(1, k).Style.Fill.BackgroundColor = XLColor.LightGray;
                workSheet.Cell(1, k).Style.Font.SetBold();
                workSheet.Cell(1, k).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
				k++;
            }

            int i = 2;
            foreach (var row in data)
            {
                int j = 1;
                foreach (PropertyInfo property in properties)
                {
                    var cell = property.GetValue(row);
                    if (cell != null)
                    {
                       //if, else blocks can be added for writing data to excel cells in desired formats. 
					   workSheet.Cell(i, j).Value = cell.ToString();  
                    }
                    j++;
                }
                i++;
            }

            int t = 1;

            foreach (PropertyInfo property in properties)
            {
                workSheet.Column(t).AdjustToContents();
                t++;
            }

            using var memoryStream = new MemoryStream();
            workBook.SaveAs(memoryStream);

            return memoryStream.ToArray();
        }

		public static bool ImportExcelFileToDb<T>(IDbFactory MyDbFactory, byte[] FileContent)
		{
			using var memoryStream = new MemoryStream(FileContent);
			using IXLWorkbook workbook = new XLWorkbook(memoryStream);

			IXLWorksheets worksheets = workbook.Worksheets;
			//If File contain multiple or zero sheets, return -1;
			if (worksheets.Count != 1)
			{
				return false;
			}

			IXLWorksheet worksheet = worksheets.First();
			int columnCount = CountColumns(worksheet);

			IXLRow _1stRow = worksheet.Rows().First();

			//<ExcelIndex,PropName>KeyValue Pair
			Dictionary<int,string> propsExcelIndexesAndNames = new();
			int j = 1;
			while (j < columnCount)
			{
				string val = _1stRow.Cell(j).Value.ToString()?.Replace(" ", "").Trim() ?? "";

				propsExcelIndexesAndNames.Add(j,val);
				j++;
			}


			//Get Property List of DataModel
			Type classType = typeof(T);
			PropertyInfo[] properties = classType.GetProperties();

			//Asign prop names to hashSet
			HashSet<string> propNamesOfDataModel = new HashSet<string>();
			foreach (PropertyInfo prop in properties)
			{
				propNamesOfDataModel.Add(prop.Name);
			}


			//Check if the given column name exists in given model. If there is a column that doesnt match with props of model, return error.
			foreach (KeyValuePair<int, string> PropExcelIndexName in propsExcelIndexesAndNames)
			{
				if(!propNamesOfDataModel.Contains(PropExcelIndexName.Value))
				{
					return false;
				}
			}

			List<T> dataToBeinserted = new();

			//Set data to object, first row contains prop names therefore skip it.
			foreach (var row in worksheet.Rows().Skip(1))
			{
				object? instance = Activator.CreateInstance(classType);

				foreach (KeyValuePair<int, string> PropNameExcelIndex in propsExcelIndexesAndNames)
				{
					string valueReadFromExcel = row.Cell(PropNameExcelIndex.Key).Value.ToString();

					PropertyInfo? prop = classType.GetProperty(PropNameExcelIndex.Value);
					Type propertyType = prop.PropertyType;

					object? convertedValue = null;
					if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
					{
						Type? underlyingType = Nullable.GetUnderlyingType(propertyType);
						convertedValue = string.IsNullOrEmpty(valueReadFromExcel) ? null : Convert.ChangeType(valueReadFromExcel, underlyingType);
					}
					else
					{
						convertedValue = Convert.ChangeType(valueReadFromExcel, propertyType);
					}
					prop?.SetValue(instance, convertedValue);
				}

				
				if (instance == null)
					throw new Exception("A null instance attempted to be inserted to list.");

				dataToBeinserted.Add((T)instance);
			}

			//insert dataToBeinserted to db.
			using var dbConnection = MyDbFactory.OpenDbConnection();

			dbConnection.InsertAll(dataToBeinserted);
			dbConnection.Close();

			return true;
		}

		// Return column count of uploaded excel.
		private int CountColumns(IXLWorksheet worksheet)
		{
			IXLRow Row1st = worksheet.Rows().First();

			int detectedColumnNumber = 0;
			int i = 1;
			//Count how many cells in first row are filled
			while (Row1st.Cell(i).Value.ToString() != "")
			{
				i++;
				detectedColumnNumber++;
			}

			return detectedColumnNumber;
		}
    }
}
