using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

    internal class ExcelStyle
    {
        public string Name { get; set; }
        public string Color { get; set; }
        public override string ToString()
        {
            return $""""
                <ss:Style ss:ID="{Name}">
                  <ss:Interior ss:Color="#{Color}" ss:Pattern="Solid"/>
                </ss:Style>
                """";
        }
    }

