using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OLEemailRetrieve
{
    public class ResponseExecuted
    {
        public string eFirstName { get; set; }
        public string eLastName { get; set; }
        public string eResponse { get; set; }
        public string eMaterialNumber { get; set; }
        public DateTime eCreationDate { get; set; }
        public string eRequestor { get; set; }
        public string eRequestorPhoneNumber { get; set; }
        public string eRequestorEmailAddress { get; set; }
        public string eCSNotes { get; set; }
        public string ePurchaseOrderToGermany { get; set; }
        public Nullable<int> eLineNumber { get; set; }
    }
}
