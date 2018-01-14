namespace InvoiceQuery
{
    internal class InvoiceItem
    {
        public double Amount { get; set; }
        public string QuickBooksID { get; set; }
        public Item Item { get; internal set; }
    }
}