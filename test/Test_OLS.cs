#r "System.IO"
#r "Microsoft.Office.Interop.Excel"

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

string fileName = @"C:\Users\trinh\Downloads\Armstrong Project\SecurityAutoBuild\config\SecurityAutoBuild"; // Enter Model Auto Build Excel file
string myWorkbook = fileName + ".xlsx";
var excelApp = new Excel.Application();

excelApp.Visible = false;
excelApp.DisplayAlerts = false;
Excel.Workbook wb = excelApp.Workbooks.Open(myWorkbook);

string[] tabs = {"Roles", "RLS", "OLS"};
int tabCount = tabs.Count();

// Default to None permission for all fact tables
foreach(var role in Model.Roles) {
	foreach(var table in Model.Tables.Where(t => t.Name.ToLower().StartsWith("fct"))) {
		table.ObjectLevelSecurity[role.Name] = MetadataPermission.None;
	}
}

for (int i=0; i<tabCount; i++)
{
    string wsName = tabs[i];
    Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[wsName];
    Excel.Range xlRange = (Excel.Range)ws.UsedRange;

    int rowCount = xlRange.Rows.Count;
	int colCount = xlRange.Columns.Count;

	for (int r=2; r<=rowCount; r++)
	{
		// OLS
		if (i==2)
		{
			string roleName = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString().Replace("*",".*");
			string objectType = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString();
			string tableName = (string)(ws.Cells[r,4] as Excel.Range).Text.ToString();
			string objectName = (string)(ws.Cells[r,5] as Excel.Range).Text.ToString().Replace("*",".*");
			string ols = (string)(ws.Cells[r,6] as Excel.Range).Text.ToString();

			// 2024-11-14 QUANG - UPDATE as chosen multiple roles based on * character
			// if (!Model.Roles.Any(a => a.Name == ro))
			if (!Model.Roles.Any(item => Regex.IsMatch(item.Name, roleName)))
		    {
		    	Error("Row level security for the '"+roleName+"' role cannot be created since the role does not exist.");
		    	return;
		    }

			foreach (var item in Model.Roles.Where(item => Regex.IsMatch(item.Name, roleName))) {

				string ro = item.Name;

				if (objectType.ToLower() == "all") {
					// ObjectType as ALL then assigned Read permission for all fact tables
					foreach(var table in Model.Tables.Where(t => t.Name.ToLower().StartsWith("fct"))) {
						table.ObjectLevelSecurity[ro] = MetadataPermission.Read;
					}
				}

				else if (objectType.ToLower() == "table") {
					if (!Model.Tables.Any(a => a.Name == objectName))
					{
						Error("Object level security for the '"+ro+"' role cannot be created since the '"+objectName+"' table does not exist.");
						return;
					}
					
					// Assign read permission to user role
					if (ols.ToLower() == "read") {
						Model.Tables[objectName].ObjectLevelSecurity[ro] = MetadataPermission.Read;
					}
					else if (ols.ToLower() == "none") {
						Model.Tables[objectName].ObjectLevelSecurity[ro] = MetadataPermission.None;
					}
					else {
						Error("Object level security for the '"+ro+"' role cannot be created since the '"+ols+"' permission does not exist.");
						return;
					}
				}

				else if (objectType.ToLower() == "column") {
					if (!Model.Tables.Any(a => a.Name == tableName))
					{
						Error("Object level security for the '"+ro+"' role cannot be created since the '"+tableName+"' table does not exist.");
						return;
					}

					// 2024-11-14 QUANG - UPDATE as chosen multiple columns based on * character
					if (!Model.Tables[tableName].Columns.Any(a => Regex.IsMatch(a.Name, objectName)))
					{
						Error("Object level security for the '"+ro+"' role cannot be created since the '"+tableName+" - "+objectName+"' column does not exist.");
						return;
					}

					// Assign read permission to user role
					if (ols.ToLower() == "read") {
						// 2024-11-14 QUANG - UPDATE as chosen multiple columns based on * character
						foreach (var c in Model.Tables[tableName].Columns.Where(a => Regex.IsMatch(a.Name, objectName))) {
							Model.Tables[tableName].Columns[c.Name].ObjectLevelSecurity[ro] = MetadataPermission.Read;
						}
					}
					else if (ols.ToLower() == "none") {
						// 2024-11-14 QUANG - UPDATE as chosen multiple columns based on * character
						foreach (var c in Model.Tables[tableName].Columns.Where(a => Regex.IsMatch(a.Name, objectName))) {
							Model.Tables[tableName].Columns[c.Name].ObjectLevelSecurity[ro] = MetadataPermission.None;
						}
					}
					else {
						Error("Object level security for the '"+ro+"' role cannot be created since the '"+ols+"' permission does not exist.");
						return;
					}
				}

				else if (objectType.ToLower() == "measure") {
					if (!Model.AllMeasures.Any(a => a.Name == objectName)) {
						Error("Object level security for the '"+ro+"' role cannot be created since the '"+objectName+"' measure does not exist.");
						return;
					}

					foreach(var m in Model.AllMeasures.Where(m => m.Name ==  objectName)) {
						foreach(var column in m.DependsOn.Columns) {
							// Assign read permission to user role
							if (ols.ToLower() == "read") {
								Model.Tables[column.DaxTableName.Replace("'", "")].Columns[column.Name].ObjectLevelSecurity[ro] = MetadataPermission.Read;
							}
							else if (ols.ToLower() == "none") {
								Model.Tables[column.DaxTableName.Replace("'", "")].Columns[column.Name].ObjectLevelSecurity[ro] = MetadataPermission.None;
							}
							else {
								Error("Object level security for the '"+ro+"' role cannot be created since the '"+ols+"' permission does not exist.");
								return;
							}
						}
					}
				}
			}
		}
	}
}

wb.Close();
excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

Info("The Role / RLS / OLS has been generated. Please check and save your model.");