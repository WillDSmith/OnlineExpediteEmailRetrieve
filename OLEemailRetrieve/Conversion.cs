using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using Microsoft.Exchange.WebServices.Data;

namespace OLEemailRetrieve
{
    class Conversion
    {

        public void ConvertXls(string fileName)
        { 
            DataTable dt = new DataTable();
            InternalEntities db = new InternalEntities();

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, false))
            {

                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (Cell cell in rows.ElementAt(0))
                {
                    dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (Row row in rows) // Include the header row
                {
                    DataRow tempRow = dt.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }

                    dt.Rows.Add(tempRow);
                }

            }
            dt.Rows.RemoveAt(0); // Remove header row.

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                var RequestId = Convert.ToInt32(dr["Request Id"]);
                var Response = dr["Response"].ToString();
                var GermanyResponder = dr["Germany Responder"].ToString();
                var resp = db.OnlineExpedites.Where(x => x.RequestId == RequestId).FirstOrDefault<OnlineExpedite>();

                if (resp != null)
                {
                    //FileStream filestream = new FileStream(log + logfilename, FileMode.Create);
                    //var streamwriter = new StreamWriter(filestream);
                    //streamwriter.AutoFlush = true;
                    //Console.SetOut(streamwriter);
                    //Console.SetError(streamwriter);
                    
                    resp.Response = Response;

                    // Get the Customer Service email and send an email with the response from Germany
                    var csEmail = resp.Email;
                    var csName = resp.Name;
                    var csDateCreated = resp.CreationDate;

                    string mEmailTo = csEmail;                                                  //ConfigurationManager.AppSettings["EmailTo"].ToString().Split(',');
                    string mEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
                    string mEmailSubject = "Reposnse from Germany, for Online Expedites";       //ConfigurationManager.AppSettings["EmailSubject"];

                    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                    service.UseDefaultCredentials = true;
                    service.AutodiscoverUrl(mEmailFrom, RedirectionUrlValidationCallback);

                    EmailMessage email = new EmailMessage(service);
                    
                    email.Subject = mEmailSubject;
                    email.Body = new MessageBody(" Hi, " + csName + " Here is the response for the request you sent on " + csDateCreated + " (" + Response + ")" );
                    //email.Attachments.AddFileAttachment(path + filename);

                    //streamwriter.WriteLine("Sending email...");
                    email.Send();
                    //streamwriter.WriteLine("Email Sent....!");
                }

                using (var dbCtx = new InternalEntities())
                {
                    dbCtx.Entry(resp).State = System.Data.Entity.EntityState.Modified;

                    dbCtx.SaveChanges();
                }


            }
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

    }
}
