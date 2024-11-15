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

string[] dimension = {"OBU"};
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
		if (i==1)
		{
			string ro = (string)(ws.Cells[r,1] as Excel.Range).Text.ToString();
			string dimensionKey = (string)(ws.Cells[r,2] as Excel.Range).Text.ToString();
			string dimensionValue = (string)(ws.Cells[r,3] as Excel.Range).Text.ToString();

		    if (!dimensionKey.Equals("OBU")) {
		    	Info("Row level security for the dimension'"+dimensionKey+"' cannot be created since the '"+dimensionKey+"' does not config.");
		    	return;
		    }

            string tableName = "DimOBU";
            string rls = "'"+tableName+"'["+dimensionKey.ToLower()+"]=\""+dimensionValue+"\"";
            int rlsLength = rls.Length;

            if (!dimensionValue.Equals("HQ")) {
                if (rls[0] == '"') {
                    rls = rls.Substring(1,rlsLength - 2);
                }
                
                rls = rls.Replace("\"\"","\"");    
                
                Model.Tables[tableName].RowLevelSecurity[ro] = rls;
            }
		}
	}
}

wb.Close();
excelApp.Quit();
System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

Info("The Role / RLS / OLS has been generated. Please check and save your model.");