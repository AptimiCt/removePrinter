using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace remove_printer
{
    class ListOfPrinter : List<string>
    {
        string namePrinter;
        string portName;

        public ListOfPrinter(string namePrinter, string portName)
        {
            this.NamePrinter = namePrinter;
            this.PortName = portName;
        }

        public string NamePrinter { get => namePrinter; set => namePrinter = value; }
        public string PortName { get => portName; set => portName = value; }

        public override string ToString()
        {
            return $"{NamePrinter}, {PortName}";
        }
    }
}
