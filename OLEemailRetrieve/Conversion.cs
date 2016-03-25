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
using System.IO;

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

            #region Email To Individual Reps
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                var RequestId = Convert.ToInt32(dr["Request Id"]);
                var Response = dr["Response"].ToString();
                var GermanyResponder = dr["Germany Responder"].ToString();
                var resp = db.OnlineExpedites.Where(x => x.RequestId == RequestId).FirstOrDefault<OnlineExpedite>();

                if (resp != null)
                {
                    // Get the Customer Service email and send an email with the response from Germany
                    resp.Response = Response;
                    var csEmail = resp.Email;   //Make sure to change this back to resp.Email
                    var csFirstName = resp.CSFirstName;
                    var csLastName = resp.CSLastName;
                    var csCreationDate = resp.CreationDate;
                    var csEDPToolNumber = resp.EDPToolNumber;
                    var csName = csFirstName + " " + csLastName;

                    var msg = csName + ",<br/>" +
                        "<br/>Here's the response for the requests you sent on " + csCreationDate + ", click <a href=\"http://staging.guhring.com/CustomerService/OnlineExpedite\">here</a>, to go to the Online Expedites Page." +
                        "<br/><br/><b>Material Number:</b> " + "<font color=\"red\">" + csEDPToolNumber + "</font>" +
                        "<br/><b>Response:</b> " + "<font color=\"red\">" + resp.Response + "</font>";


                    string log = @"CSLogs\";
                    string dtt = DateTime.Now.ToString("MMddyyyy");
                    string logfilename = csName + dtt + ".txt";

                    FileStream filestream = new FileStream(log + logfilename, FileMode.Create);
                    var streamwriter = new StreamWriter(filestream);
                    streamwriter.AutoFlush = true;
                    Console.SetOut(streamwriter);
                    Console.SetError(streamwriter);

                    string mEmailTo = csName;                                                  //ConfigurationManager.AppSettings["EmailTo"].ToString().Split(',');
                    string mEmailFrom = "wilsmi@guhring.com";                                   //ConfigurationManager.AppSettings["EmailFrom"];
                    string mEmailSubject = "Response from Germany, for Online Expedites";       //ConfigurationManager.AppSettings["EmailSubject"];

                    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                    service.UseDefaultCredentials = true;
                    service.AutodiscoverUrl(mEmailFrom, RedirectionUrlValidationCallback);

                    EmailMessage message = new EmailMessage(service);

                    message.ToRecipients.Add(csEmail);
                    message.Subject = mEmailSubject;
                    message.Body = new MessageBody();
                    message.Body.BodyType = BodyType.HTML;
                    message.Body = msg;

                    streamwriter.WriteLine("Sending email...");
                    message.Send();
                    streamwriter.WriteLine("Email Sent....!");
                }

                using (var dbCtx = new InternalEntities())
                {
                    dbCtx.Entry(resp).State = System.Data.Entity.EntityState.Modified;

                    dbCtx.SaveChanges();
                }
            }
            #endregion

            #region Email TO Administrators
            List<ResponseExecuted> listResponseExecuted = new List<ResponseExecuted>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                
                var RequestId = Convert.ToInt32(dr["Request Id"]);
                var Response = dr["Response"].ToString();
                var GermanyResponder = dr["Germany Responder"].ToString();
                var resp = db.OnlineExpedites.Where(x => x.RequestId == RequestId).FirstOrDefault<OnlineExpedite>();

                if (resp != null)
                {
                    // Get the Customer Service email and send an email with the response from Germany
                    var data = new ResponseExecuted
                    {
                        eResponse = Response,
                        eFirstName = resp.CSFirstName,
                        eLastName = resp.CSLastName,
                        eCreationDate = Convert.ToDateTime(resp.CreationDate),
                        eMaterialNumber = resp.EDPToolNumber
                    };

                    listResponseExecuted.Add(data);
                }
            }

            List<string> lstmsgs = new List<string>(); 
           
            //string mEmailTo = csName; 
            String[] aEmailTo = ConfigurationManager.AppSettings["AdminsTo"].ToString().Split(',');
            string aEmailFrom = ConfigurationManager.AppSettings["AdminsFrom"];
            string aEmailSubject = ConfigurationManager.AppSettings["AdminsEmailSubject"];
            string greeting = "";
            string lstOfResponses = "";

            ExchangeService aService = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            aService.UseDefaultCredentials = true;
            aService.AutodiscoverUrl(aEmailFrom, RedirectionUrlValidationCallback);

            EmailMessage aMessage = new EmailMessage(aService);

            foreach (String s in aEmailTo)
            {
                aMessage.ToRecipients.Add(s);
            }
            aMessage.Subject = aEmailSubject;
            aMessage.Body = new MessageBody();
            aMessage.Body.BodyType = BodyType.HTML;
            
            foreach (var r in listResponseExecuted)
            {
                greeting = "Greetings, <br/>" +
                        "<br/>Here's the responses for the requests sent on " + r.eCreationDate + ", click <a href=\"http://staging.guhring.com/CustomerService/OnlineExpedite\">here</a>, to go to the Online Expedites Page.";
                string msg = 
                        "<br/><br/><b>Customer Service Rep:</b> " + "<font color=\"red\">" + r.eFirstName + " " + r.eLastName + "</font>" +
                        "<br/><b>Material Number:</b> " + "<font color=\"red\">" + r.eMaterialNumber + "</font>" +
                        "<br/><b>Response:</b> " + "<font color=\"red\">" + r.eResponse + "</font>";
                lstmsgs.Add(msg);
            }
           
            var sb = new StringBuilder();
            lstmsgs.ForEach(s => sb.Append(s));
            lstOfResponses = sb.ToString();

            aMessage.Body = greeting + lstOfResponses;
            aMessage.Send();
           
            #endregion  
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
