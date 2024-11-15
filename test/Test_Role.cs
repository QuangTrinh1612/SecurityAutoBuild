#r "System.IO"
#r "Microsoft.Office.Interop.Excel"

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

string fileName = @"C:\Users\trinh\Downloads\Armstrong Project\SecurityAutoBuild\config\SecurityAutoBuild"; // Enter Model Auto Build Excel file
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
	}
}

wb.Close();
excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

Info("The Role / RLS / OLS has been generated. Please check and save your model.");