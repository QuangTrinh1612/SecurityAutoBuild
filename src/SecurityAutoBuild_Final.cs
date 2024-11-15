#r "System.IO"
#r "Microsoft.Office.Interop.Excel"

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

public class SecurityModel {
    public static void Main() {
        string fileName = @"C:\Users\trinh\Downloads\Armstrong Project\SecurityAutoBuild\config\SecurityAutoBuild"; // Enter Model Auto Build Excel file
        
        addRole(fileName);
    }
    
    public static string addRole(string fileName) {
        
    }
}

SecurityModel.Main();