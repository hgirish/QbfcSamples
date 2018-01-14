namespace InvoiceQuery
{
    public class InvoiceItem
    {
        public double? Amount { get; set; }
        public string QuickBooksID { get; set; }
        public Item Item { get; internal set; }
        public string Description { get; internal set; }
    }
}