# FullyGenericExcelExportImport
	This class allows you to: 
	Export a database query result into an excel file 
	and
	Import data of excel file into database table 
	in a completely generic way.
	
	Import: First line of excel file must contain column names of database table and backend data model. 
	Orders of columns in excel file is not important, also missing columns will be tried to inserted as null if database allows.
	First empty cell in the first row will be accepted as last column.
	
	Export: Just call the function with any data model.
