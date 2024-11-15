#r "System.IO"

using System.IO;
using System.Data;

string roleTableName = "PBI_ROLE";
string rlsTableName = "PBI_RLS";
string olsTableName = "PBI_OLS";

for (int i=0; i<3; i++)
{
	// RLS
	if (i==1)
	{
		string dax = "EVALUATE '"+rlsTableName+"'";

		foreach(System.Data.DataTable table in ExecuteDax(dax).Tables) {
			
			int rowCount = table.Rows.Count;
			
			for (int r=0; r<rowCount; r++) {
				
				string ro = (string)(table.Rows[r][0]);
				string tableName = (string)(table.Rows[r][1]);
				string rls = (string)(table.Rows[r][2]);
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
		}
	}
}

Info("The Role / RLS / OLS has been executed. Please check and save your model.");