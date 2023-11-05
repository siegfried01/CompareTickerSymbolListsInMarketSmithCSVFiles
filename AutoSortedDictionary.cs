using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace CompareTickerSymbolListsInCSVFiles
{

    public class AutoInitSortedDictionary<K, V> : System.Collections.Generic.SortedDictionary<K, V> where K : notnull where V : class?
    {
        public new V this[K k]
        {
            get
            {
                if (base.Keys.Contains(k))
                    return base[k];
                else
                {
                    V v = null;
                    base.Add(k, v);
                    return v;
                }
            }
            set
            {
                base[k] = value;
            }
        }
    }
    public class AutoInitIntSortedDictionary<K> : System.Collections.Generic.SortedDictionary<K, int> where K : notnull
    {
        public new int this[K k]
        {
            get
            {
                if (base.Keys.Contains(k))
                    return base[k];
                else
                {
                    int v = 0;
                    base.Add(k, v);
                    return v;
                }
            }
            set
            {
                base[k] = value;
            }
        }
    }
    public class AutoMultiDimSortedDictionary<K, V> : System.Collections.Generic.SortedDictionary<K, V> where V : new() where K : notnull
    {
        public new V this[K k]
        {
            get
            {
                if (base.Keys.Contains(k))
                    return base[k];
                else
                {
                    var v = new V();
                    base.Add(k, v);
                    return v;
                }
            }
            set { base[k] = value; }
        }
    }
    /*
    public class AutoMultiDimDescendingSortedDictionary<K, V> : System.Collections.Generic.SortedDictionary<K, V> where V : new() where K : notnull
    {
        // Severity	Code	Description	Project	File	Line	Suppression State
        // Error CS0314  The type 'K' cannot be used as type parameter 'T' in the generic type or method 'DescendingComparer<T>'. There is no boxing conversion or type parameter conversion from 'K' to 'System.IComparable<K>'.	CompareTickerSymbolListsInCSVFiles c:\Users\shein\Documents\WinOOP\C#\CompareTickerSymbolListsInCSVFiles\AutoSortedDictionary.cs	72	Active

        public AutoMultiDimDescendingSortedDictionary() : base(comparer: new DescendingComparer< *K* >())  {  }
        public new V this[K k]
        {
            get
            {
                if (base.Keys.Contains(k))
                    return base[k];
                else
                {
                    var v = new V();
                    base.Add(k, v);
                    return v;
                }
            }
            set { base[k] = value; }
        }
    }
    */
}
