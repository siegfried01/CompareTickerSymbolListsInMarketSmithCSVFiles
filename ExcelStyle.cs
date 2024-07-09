using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

    internal class ExcelStyle
    {
        public string Name { get; set; }
        public string Color { get; set; }
        public string Pattern { get; set; } = "Solid";
        public string PatternColor { get; set; } = "";
        public override string ToString()
        {
            var patColor = string.IsNullOrEmpty(PatternColor)?"":$"""ss:PatternColor="#{PatternColor}" """;
            var result = $""""
                <ss:Style ss:ID="{Name}">
                  <ss:Interior ss:Color="#{Color}" ss:Pattern="{Pattern}" {patColor}/>
                </ss:Style>
                """";
            return result;
        }
    }

