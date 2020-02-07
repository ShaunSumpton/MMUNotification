using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.Office.Interop.Outlook;
using System.Data.SqlClient;
using System.Data;
using System.Collections;
using DataTable = System.Data.DataTable;
using System.Data.OleDb;

namespace MMUNotification
{
    class Program
    {
        static void Main(string[] args)
        {

            string[] filePaths = Directory.GetFiles(@"\\6.1.1.37\SFTPRoot\Manchester Metropolitan University", "*.xlsx"); // find worksheet on SFTP

            Console.WriteLine(filePaths[0]);

            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application(); // Create Excel Instance

            Workbook exceldoc = application.Workbooks.Open(filePaths[0]); // create workbook
            Worksheet ws; // create worksheet

            ws = (Worksheet)exceldoc.Worksheets[1]; // worksheet assigned to 1st sheet in workbook


            int LastRow = ws.UsedRange.Rows.Count;    // find last row and last column of sheet
            _ = ws.UsedRange.Columns.Count;
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            _ = ws.get_Range("A1", last);
            Range uknot = ws.Columns["Q"]; // column to count UK or NON UK sends
            _ = last.Row;

            var UK = application.WorksheetFunction.CountIf(uknot, "UK"); // count uk sends
            var NONUK = application.WorksheetFunction.CountIf(uknot, "Non-UK"); // count nonuk sends

            Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application(); // create outlook instance
            MailItem mailItem = app.CreateItem(OlItemType.olMailItem); // create mail item




            mailItem.Subject = "MMU Data Notification " + DateTime.Now.ToString("dd/MM/yyyy");                                                    // set up email with to,subject, body etc
            mailItem.To = "s.sumpton@agnortheast.com; S.kent@agnortheast.com ; r.richardson@agnortheast.com";


            mailItem.Importance = OlImportance.olImportanceHigh;
            mailItem.Display(false); // dont display mail item before sending
            _ = mailItem.HTMLBody;
            var body = "MMU Offer Guide Quantities <br /> <br />" + "Number of UK: " + UK + "<br />" + "Number of Non-UK: " + NONUK + "<br />" + "--------------------------------------";
            mailItem.HTMLBody = body; //+ signature;
           mailItem.Send(); // senmail confirming data count

        }
    }
}
