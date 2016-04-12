using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Configuration;
using Microsoft.Exchange.WebServices.Data;

namespace OLEemailRetrieve
{
    internal class Conversion
    {

        public void ConvertXls(string fileName)
        {
            DataTable dt = new DataTable();
            var db = new InternalEntities();

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                var sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                var relationshipId = sheets.First().Id.Value;
                var worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                var workSheet = worksheetPart.Worksheet;
                var sheetData = workSheet.GetFirstChild<SheetData>();
                var rows = sheetData.Descendants<Row>();

                var enumerable = rows as Row[] ?? rows.ToArray();
                foreach (var cell in enumerable.ElementAt(0).Cast<Cell>())
                {
                    dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (Row row in enumerable) // Include the header row
                {
                    var tempRow = dt.NewRow();

                    //for (int i = 0; i < row.Descendants<Cell>().Count(); i++)  // Was previously this, changed to dt.Columns.count
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }
                    dt.Rows.Add(tempRow);
                }
            }
            dt.Rows.RemoveAt(0); // Remove header row.

            #region Email To Individual Reps
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                var requestId = Convert.ToInt32(dr["Request Id"]);
                var response = dr["Response"].ToString();
                var germanyResponder = dr["Germany Responder"].ToString();
                var resp = db.OnlineExpedites.FirstOrDefault(x => x.RequestId == requestId);

                if (resp != null)
                {
                    // Get the Customer Service email and send an email with the response from Germany
                    resp.Response = response;
                    resp.GermanyResponder = germanyResponder;
                    var csEmail = resp.Email;   //Make sure to change this back to resp.Email
                    var csFirstName = resp.CSFirstName;
                    var csLastName = resp.CSLastName;
                    var csCreationDate = resp.CreationDate;
                    var csEdpToolNumber = resp.EDPToolNumber;
                    var csRequestor = resp.Requestor;
                    var csRequestorPhone = resp.RequestorPhoneNumber;
                    var csResquestorEmail = resp.RequestorEmailAddress;
                    var csCsNotes = resp.CSNotes;
                    var csName = csFirstName + " " + csLastName;

                    var msg = csName + ",<br/>" +
                        "<br/>Here's the response for the requests you sent on " + csCreationDate + ", click <a href=\"http://staging.guhring.com/CustomerService/OnlineExpedite\">here</a>, to go to the Online Expedites Page." +
                        "<br/><br/><b>Material Number:</b> " + "<font color=\"red\">" + csEdpToolNumber + "</font>" +
                        "<br/><b>Response:</b> " + "<font color=\"red\">" + resp.Response + "</font>" +
                        "<br/><b>Requestor:</b> " + "<font color=\"red\">" + csRequestor + "</font>" +
                        "<br/><b>Requestor Phone Number:</b> " + "<font color=\"red\">" + csRequestorPhone + "</font>" +
                        "<br/><b>Requestor Email:</b> " + "<font color=\"red\">" + csResquestorEmail + "</font>" +
                        "<br/><b>CS Notes:</b> " + "<font color=\"red\">" + csCsNotes + "</font>";
 
                    string mEmailFrom    = ConfigurationManager.AppSettings["EmailFrom"];
                    string mEmailSubject = ConfigurationManager.AppSettings["EmailSubject"];

                    ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                    
                    if (ConfigurationManager.AppSettings["UseDefault"] != "true")
                    {
                        service.Credentials = new WebCredentials(ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Password"], ConfigurationManager.AppSettings["Domain"]);
                        service.TraceEnabled = true;
                        service.TraceFlags = TraceFlags.All;

                        try
                        {
                            service.AutodiscoverUrl(ConfigurationManager.AppSettings["Address"], RedirectionUrlValidationCallback);
                        }
                        catch
                        {
                            Console.WriteLine("Not a valid [Address] in app config");
                        }

                    }
                    else
                    {
                        service.UseDefaultCredentials = true;
                        try
                        {
                            service.AutodiscoverUrl(mEmailFrom, RedirectionUrlValidationCallback);
                        }
                        catch
                        {
                            Console.WriteLine("Not a valid [EMAILFROM] in app config");
                        }

                    }

                    EmailMessage message = new EmailMessage(service);

                    message.ToRecipients.Add(csEmail);
                    message.Subject = mEmailSubject;
                    message.Body = new MessageBody {BodyType = BodyType.HTML};
                    message.Body = msg;

                    Console.WriteLine("Sending email...");
                    message.Send();
                    Console.WriteLine("Email Sent....!");
                }

                using (var dbCtx = new InternalEntities())
                {
                    dbCtx.Entry(resp).State = System.Data.Entity.EntityState.Modified;

                    dbCtx.SaveChanges();
                }
            }
            #endregion

            #region Email TO Administrators
            var listResponseExecuted = new List<ResponseExecuted>();
            for (var i = 0; i < dt.Rows.Count; i++)
            {
                DataRow dr = dt.Rows[i];
                
                var requestId = Convert.ToInt32(dr["Request Id"]);
                var response = dr["Response"].ToString();
                var germanyResponder = dr["Germany Responder"].ToString();
                var resp = db.OnlineExpedites.FirstOrDefault(x => x.RequestId == requestId);

                if (resp == null) continue;
                resp.GermanyResponder = germanyResponder;
                    
                // Get the Customer Service email and send an email with the response from Germany
                var data = new ResponseExecuted
                {
                    eResponse = response,
                    eFirstName = resp.CSFirstName,
                    eLastName = resp.CSLastName,
                    eCreationDate = Convert.ToDateTime(resp.CreationDate),
                    eMaterialNumber = resp.EDPToolNumber,
                    eRequestor = resp.Requestor,
                    eRequestorPhoneNumber = resp.RequestorPhoneNumber,
                    eRequestorEmailAddress = resp.RequestorEmailAddress,
                    eCSNotes = resp.CSNotes,
                    ePurchaseOrderToGermany = resp.PurchaseOrderToGermany,
                    eLineNumber = resp.LineNumber
                };

                listResponseExecuted.Add(data);
            }

            var lstmsgs = new List<string>(); 
            
            String[] aEmailTo = ConfigurationManager.AppSettings["AdminsTo"].Split(',');
            var aEmailFrom = ConfigurationManager.AppSettings["AdminsFrom"];
            var aEmailSubject = ConfigurationManager.AppSettings["AdminsEmailSubject"];
            var greeting = "";

            ExchangeService aService = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            

            if (ConfigurationManager.AppSettings["UseDefault"] != "true")
            {
                aService.Credentials = new WebCredentials(ConfigurationManager.AppSettings["User"], ConfigurationManager.AppSettings["Password"], ConfigurationManager.AppSettings["Domain"]);
                aService.TraceEnabled = true;
                aService.TraceFlags = TraceFlags.All;

                try
                {
                    aService.AutodiscoverUrl(ConfigurationManager.AppSettings["Address"], RedirectionUrlValidationCallback);
                }
                catch
                {
                    Console.WriteLine("Not a valid [Address] in app config");
                }

            }
            else
            {
                aService.UseDefaultCredentials = true;
                try
                {
                    aService.AutodiscoverUrl(aEmailFrom, RedirectionUrlValidationCallback);
                }
                catch
                {
                    Console.WriteLine("Not a valid [EMAILFROM] in app config");
                }

            }

            var aMessage = new EmailMessage(aService);

            foreach (String s in aEmailTo)
            {
                aMessage.ToRecipients.Add(s);
            }
            aMessage.Subject = aEmailSubject;
            aMessage.Body = new MessageBody {BodyType = BodyType.HTML};

            foreach (var r in listResponseExecuted)
            {
                greeting = "Greetings, <br/>" +
                        "<br/>Here's the responses for the requests sent on " + r.eCreationDate + ", click <a href=\"http://staging.guhring.com/CustomerService/OnlineExpedite\">here</a>, to go to the Online Expedites Page.";
                var msg = 
                        "<br/><br/><b>Customer Service Rep:</b> " + "<font color=\"red\">" + r.eFirstName + " " + r.eLastName + "</font>" +
                        "<br/><b>Material Number:</b> " + "<font color=\"red\">" + r.eMaterialNumber + "</font>" +
                        "<br/><b>Response:</b> " + "<font color=\"red\">" + r.eResponse + "</font>" +
                        "<br/><b>Requestor:</b> " + "<font color=\"red\">" + r.eRequestor + "</font>" +
                        "<br/><b>Requestor Phone Number:</b> " + "<font color=\"red\">" + r.eRequestorPhoneNumber + "</font>" +
                        "<br/><b>Requestor Email:</b> " + "<font color=\"red\">" + r.eRequestorEmailAddress + "</font>" +
                        "<br/><b>CS Noted:</b> " + "<font color=\"red\">" + r.eCSNotes + "</font>" +
                        "<br/><b>Line Number:</b> " + "<font color=\"red\">" + r.eLineNumber + "</font>" +
                        "<br/><b>P.O. To Germany:</b> " + "<font color=\"red\">" + r.ePurchaseOrderToGermany + "</font>";
                lstmsgs.Add(msg);
            }
           
            var sb = new StringBuilder();
            lstmsgs.ForEach(s => sb.Append(s));
            var lstOfResponses = sb.ToString();

            aMessage.Body = greeting + lstOfResponses;
            aMessage.Send();
           
            #endregion  
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            var value = cell.CellValue.InnerXml;

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
