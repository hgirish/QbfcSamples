using System.Collections.Generic;

namespace InvoiceQuery
{
    public class Invoice
    {
        public string QuickBooksID { get; internal set; }
        public string EditSequence { get; internal set; }
        public Customer Customer { get; internal set; }
        public IList<InvoiceItem> InvoiceItems { get; set; } = new List<InvoiceItem>();
        public string InvoiceNumber { get; internal set; }
        public string Memo { get; internal set; }
        public string JobNumber { get; internal set; }
        public string CustomerName { get; internal set; }
    }
}