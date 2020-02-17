using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Net.Mail;

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

            //r.Richardson@agnortheast.com

            MailMessage message = new MailMessage("s2@agnortheast.com", "S.Sumpton@agnortheast.com; s.kent@agnortheast.com; a.granger@agnortheast.com",
            "MMU Data Notification " + DateTime.Now.ToString("dd/MM/yyyy"),
                "MMU Offer Guide Quantities <br /> <br />" + "Number of UK: " + UK + "<br />" + "Number of Non-UK: " + NONUK + "<br />" + "--------------------------------------");
            message.IsBodyHtml = true;
            SmtpClient client = new SmtpClient("6.1.1.143");
            client.Send(message);


        }
    }
}


