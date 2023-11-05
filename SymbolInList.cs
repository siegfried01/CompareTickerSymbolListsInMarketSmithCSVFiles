using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompareTickerSymbolListsInCSVFiles;

	class SymbolInList
	{
		public string Name { get; set; } = "";
		public SortedSet<string> Lists { get; set; } = new SortedSet<string>();
	}
