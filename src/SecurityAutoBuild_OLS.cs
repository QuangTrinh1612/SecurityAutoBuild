#r "System.IO"

using System.IO;
using System.Data;

string roleTableName = "PBI_ROLE";
string rlsTableName = "PBI_RLS";
string olsTableName = "PBI_OLS";

for (int i=0; i<3; i++)
{
	if (i==2) {
		string dax = "EVALUATE '"+olsTableName+"'";

		foreach(System.Data.DataTable table in ExecuteDax(dax).Tables) {
			
			int rowCount = table.Rows.Count;
			
			for (int r=0; r<rowCount; r++) {
				string ro = (string)(table.Rows[r][1]);
				string objectType = (string)(table.Rows[r][2]);
				// string tableName = (string)(table.Rows[r][3]);
				string tableName = table.Rows[r][3] != DBNull.Value ? (string)table.Rows[r][3] : string.Empty;
				string objectName = (string)(table.Rows[r][4]);
				string ols = (string)(table.Rows[r][5]);

				if (!Model.Roles.Any(a => a.Name == ro))
				{
					Error("Object level security for the '"+ro+"' role cannot be created since the role does not exist.");
					return;
				}

				if (objectType == "Table") {
					if (!Model.Tables.Any(a => a.Name == objectName))
					{
						Error("Object level security for the '"+ro+"' role cannot be created since the '"+objectName+"' table does not exist.");
						return;
					}
					
					// Assign read permission to user role
					if (ols == "Read") {
						Model.Tables[objectName].ObjectLevelSecurity[ro] = MetadataPermission.Read;
					}
					else if (ols == "None") {
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
					if (ols == "Read") {
						Model.Tables[tableName].Columns[objectName].ObjectLevelSecurity[ro] = MetadataPermission.Read;
					}
					else if (ols == "None") {
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
							if (ols == "Read") {
								Model.Tables[column.DaxTableName.Replace("'", "")].Columns[column.Name].ObjectLevelSecurity[ro] = MetadataPermission.Read;
							}
							else if (ols == "None") {
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

Info("The Role / RLS / OLS has been executed. Please check and save your model.");