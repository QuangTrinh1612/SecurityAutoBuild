#r "System.IO"
#r "Microsoft.Office.Interop.Excel"

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

string fileName = @"C:\Users\quan3460\Downloads\Armstrong Project\SecurityAutoBuild\SecurityAutoBuild_New"; // Enter Model Auto Build Excel file
string myWorkbook = fileName + ".xlsx";
var excelApp = new Excel.Application();

excelApp.Visible = false;
excelApp.DisplayAlerts = false;
Excel.Workbook wb = excelApp.Workbooks.Open(myWorkbook);

string[] tabs = {"Roles", "RLS", "OLS"};
int tabCount = tabs.Count();

for (int i=0; i<tabCount; i++)
{
    string wsName = tabs[i];
    Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[wsName];
    Excel.Range xlRange = (Excel.Range)ws.UsedRange;

    int rowCount = xlRange.Rows.Count;
	int colCount = xlRange.Columns.Count;

	for (int r=2; r<=rowCount; r++)
	{
		// Roles
		if (i==0)
		{
			string ro = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
			// string rm = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
			string mp = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString().ToLower();

			// Add Roles and do not duplicate
		    if (!Model.Roles.ToList().Any(x => x.Name == ro))
		    {
			    var obj = Model.AddRole(ro);
			    // obj.RoleMembers = rm;

			    if (mp.ToLower() == "read")
			    {
			        obj.ModelPermission = ModelPermission.Read;
			    }
			    else if (mp.ToLower() == "admin")
			    {
			        obj.ModelPermission = ModelPermission.Administrator;
			    }
			    else if (mp.ToLower() == "refresh")
			    {
			        obj.ModelPermission = ModelPermission.Refresh;
			    }
			    else if (mp.ToLower() == "readrefresh")
			    {
			        obj.ModelPermission = ModelPermission.ReadRefresh;
			    }
			    else if (mp.ToLower() == "none")
			    {
			        obj.ModelPermission = ModelPermission.None;
			    }
			}
		}

		// RLS
		else if (i==1)
		{
			string ro = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
			string tableName = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
			string rls = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString();
		    int rlsLength = rls.Length;

		    if (!Model.Tables.Any(a => a.Name == tableName))
		    {
		    	Error("Row level security for the '"+ro+"' role cannot be created since the '"+tableName+"' table does not exist.");
		    	return;
		    }

		    if (!Model.Roles.Any(a => a.Name == ro))
		    {
		    	Error("Row level security for the '"+ro+"' role cannot be created since the role does not exist.");
		    	return;
		    }

		    if (rls[0] == '"')
	        {
				rls = rls.Substring(1,rlsLength - 2);
	        }
		    
		    rls = rls.Replace("\"\"","\"");    
		    
		    Model.Tables[tableName].RowLevelSecurity[ro] = rls;  
		}

		// OLS
		else if (i==2)
		{
			string ro = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
			string objectType = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString();
			string tableName = (string)(ws.Cells[r,4] as Excel.Range).Text.ToString();
			string objectName = (string)(ws.Cells[r,5] as Excel.Range).Text.ToString();
			string ols = (string)(ws.Cells[r,6] as Excel.Range).Text.ToString();

			if (!Model.Roles.Any(a => a.Name == ro))
		    {
		    	Error("Row level security for the '"+ro+"' role cannot be created since the role does not exist.");
		    	return;
		    }

			if (objectType == "Table") {
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

			else if (objectType == "Column") {
				if (!Model.Tables.Any(a => a.Name == tableName))
				{
					Error("Object level security for the '"+ro+"' role cannot be created since the '"+tableName+"' table does not exist.");
					return;
				}

				if (!Model.Tables[tableName].Columns.Any(a => a.Name == objectName))
				{
					Error("Object level security for the '"+ro+"' role cannot be created since the '"+tableName+" - "+objectName+"' column does not exist.");
					return;
				}

				// Assign read permission to user role
				if (ols.ToLower() == "read") {
					Model.Tables[tableName].Columns[objectName].ObjectLevelSecurity[ro] = MetadataPermission.Read;
				}
				else if (ols.ToLower() == "none") {
					Model.Tables[tableName].Columns[objectName].ObjectLevelSecurity[ro] = MetadataPermission.None;
				}
				else {
					Error("Object level security for the '"+ro+"' role cannot be created since the '"+ols+"' permission does not exist.");
					return;
				}
			}

			else if (objectType == "Measure") {
				if (!Model.AllMeasures.Any(a => a.Name == objectName))
				{
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

wb.Close();
excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

Info("The Role / RLS / OLS has been generated. Please check and save your model.");