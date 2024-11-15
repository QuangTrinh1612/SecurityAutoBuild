#r "System.IO"

using System.IO;
using System.Data;

string roleTableName = "PBI_ROLE";
string rlsTableName = "PBI_RLS";
string olsTableName = "PBI_OLS";

for (int i=0; i<3; i++)
{
	// Roles
	if (i==0)
	{
		string dax = "EVALUATE '"+roleTableName+"'";

		foreach(System.Data.DataTable table in ExecuteDax(dax).Tables) {
			
			int rowCount = table.Rows.Count;
			
			for (int r=0; r<rowCount; r++) {
				
				string ro = (string)(table.Rows[r][0]);
				// string rm = (string)(table.Rows[r][1]);
				string mp = (string)(table.Rows[r][2]);

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
}

Info("The Role / RLS / OLS has been executed. Please check and save your model.");