using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InvoiceAddApp
{
    class Program
    {
        static void Main(string[] args)
        {
            InvoiceAddHelper addHelper = new InvoiceAddHelper();
            addHelper.DoInvoiceAdd();
        }
    }
}
