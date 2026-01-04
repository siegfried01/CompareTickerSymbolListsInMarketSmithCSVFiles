using CompareTickerSymbolListsInCSVFiles;
using CsvHelper;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using static System.Console;
using static System.Environment;
using static System.Math;
/**
 * emacs F10
 * Begin commands to execute this file using MS.NET with bash
 * if test CompareTickerSymbolListsInCSVFilesMainProgram.cs -nt "${WINOOP}/C#/CompareTickerSymbolListsInCSVFiles/bin/Debug/net9.0/CompareTickerSymbolListsInCSVFiles.exe"
 * then
 * if [ -e CompareTickerSymbolListsInCSVFilesMainProgram.exe ]
 * then
 * echo rm CompareTickerSymbolListsInCSVFilesMainProgram.exe
 * rm "${WINOOP}/C#/CompareTickerSymbolListsInCSVFiles/bin/Debug/net9.0/CompareTickerSymbolListsInCSVFiles.exe"
 * fi
 * echo devenv CompareTickerSymbolListsInCSVFiles.csproj /Build
 * devenv CompareTickerSymbolListsInCSVFiles.csproj /Build
 * else
 * echo already compiled
 * fi
 * End commands to execute this file using MS.NET with bash
 */
//
// Earngins & Sales data included. Allow override of hard coded filenames "MinDollarVol10MComp50ETF.csv" and "MinDollarVol10MComp50.csv". This works!
// ${STACKS}/Users/shein/Documents/WinOOP/C#/CompareTickerSymbolListsInCSVFiles/CompareTickerSymbolListsInCSVFilesMainProgram00090.cs

// Sample contents of formula file called "ttt SellSignalMCP Formulas.csv""
//
// Symbol,Date,Formulas
// AAPL, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// AMZN, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// ARM,  2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// AVGO, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// BKNG, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// META, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// MPLX, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// MSFT, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// NFLX, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// NVDA, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// PANW, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// PLTR, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// RACE, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// SE,   2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// SFM,  2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// SPHQ, 2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])
// TSM,  2025/03/16, =(RC[-8]-RC[-4])/(RC[-8]/(1+RC[22]/100)-RC[-4])


record AttributeAttributes(string name, 
    bool numeric, 
    bool screenTip, 
    string displayName, 
    Func<string, string> convert, 
    Func<string, string, string>? orderByConvert, 
    double max = double.MinValue, 
    double min=double.MaxValue,
    Action<CompareTickerSymbolListsInCSVFilesMainProgram, AutoInitSortedDictionary<string, string>,string>?attributeColumns=null);
class DescendingComparer<T> : IComparer<T> where T : IComparable<T>
{
    public static bool reverse { get; set; } = false;
    public int Compare(T x, T y)
    {
        if (reverse) return y.CompareTo(x);
        else return x.CompareTo(y);
    }
}
class History
{
    public DateTime Date { get; set; }
    public bool FileExists { get; set; }
    public string Name { get; set; }
}
//record DataFileCatagories(SortedSet<string> preferred, SortedSet<string> missingDataColums, SortedSet<string> donotknow);
internal class CompareTickerSymbolListsInCSVFilesMainProgram
{
    public CompareTickerSymbolListsInCSVFilesMainProgram()
    {

    }
    static SortedSet<string> preferredFiles = new SortedSet<string>() {
         "Accelerating Leaders.csv",
         //"Additions.csv",
         "All RS Line New High.csv",
         //"Deletions.csv",
         "Established Downtrend.csv",
         "Extended Stocks.csv",
         "IBD Big Cap 20.csv",
         "IBD Live Hi Closing Range.csv",
         "IBD Live Hi Weekly Closing Range.csv",
         "IBD Live Ready.csv",
         "IBD Live Watch.csv",
         "IBD 50 Index.csv",
         "LB Leaders.csv",
         "LB Sector Leaders.csv",
         "LB Spotlight.csv",
         "LB Watch.csv",
         "Large Cap.csv",
         "Long Term Leaders Watch.csv",
         "Long Term Leaders.csv",
         "ETF Indicies.csv",
         //"Top Sectors ETFs.csv",
         //"Merrill-Holdings-IRR.csv",
         //"Merrill-Holdings-Unrealized-Gain-Loss-Summary.csv",
         //"Merrill-Holdings.csv",
         "Mid Cap.csv",
         //"MinDollarVol10MComp50.csv",
         //"MinDollarVol10MComp50ETF.csv",
         "Power from Pivot.csv",
         "RS Line 5% New High.csv",
         "RS Line Blue Dot.csv",
         "RS Line New High.csv",
         "Small Cap.csv",
         //"ttt ML Holdings.csv",
         "Top Rated Stocks.csv",
         //"ETF Indicies.csv",
         //"Top Sectors ETFs.csv",
         "Near Pivot.csv",
         "Breaking Out Today.csv",
         "Watch.csv",
         //"watch closely.csv",
         "William J. O'Neil.csv",
         "Up on Volume $20M.csv",
         "Top Rated IPOs.csv",
    };
    static SortedSet<string> confidential = new SortedSet<string>(){
        "ttt ML Holdings ProfitLossDollar.csv",
        "ttt ML Holdings Shares.csv",
        "ttt ML Holdings NetLiquidValue.csv",
//        "sss Schwab Shares.csv",
//        "sss Schwab NetLiquidValue.csv",
//        "sss Schwab ProfitLossDollar.csv"
        };
    static string? envVarSkipConfidential = GetEnvironmentVariable("SKIPCONF");
    static bool skipConfidential = envVarSkipConfidential != null && envVarSkipConfidential.ToLower() == "true" ? true : false;

    /// <summary>
    ///  In some cases file names may already contain a sort order prefix. In other cases, the copy of the file name used as an index in a
    ///  dictionary will be enhanced with a sort order prefix. The sort order prefix will be removed (along with the file extension)
    ///  when converting the file name into a excel column header.
    /// </summary>
    static string patSortOrderPrefix = "aaa[0-9]{0,3}|bbb[0-9]{0,3}|ccc[0-9]{0,3}|ddd[0-9]{0,3}|eee[0-9]{0,3}|fff[0-9]{0,3}|ggg[0-9]{0,3}|hhh[0-9]{0,3}|iii[0-9]{0,3}|jjj[0-9]{0,3}|kkk[0-9]{0,3}|lll[0-9]{0,3}|mmm[0-9]{0,3}|nnn[0-9]{0,3}|ooo[0-9]{0,3}|ppp[0-9]{0,3}|qqq[0-9]{0,3}|rrr[0-9]{0,3}|sss[0-9]{0,3}|ttt[0-9]{0,3}|uuu[0-9]{0,3}|vvv[0-9]{0,3}|www[0-9]{0,3}|xxx[0-9]{0,3}|yyy[0-9]{0,3}|zzz[0-9]{0,3}";
    static Regex patFileNameOrderPrefix = new Regex(@$"^(({patSortOrderPrefix})\s*)?(.*)$"); // optional file name prefix for implementing the order of the columns
    static Regex patFileExtension = new Regex("^([^\\.]+)\\.([^\\.]+)$");
    static Regex patGotoRow = new Regex("^--GotoRow=([0-9]+)$");
    //static Regex patSymbol = new Regex(@$"({patSortOrderPrefix})([a-zA-Z0-9]*)\s*Symbol(\.csv)?$");
    //[GeneratedRegex(@"zzz([a-zA-Z0-9]*)Notes(\.csv)?$")]
    static Regex patNotes = new Regex(@$"({patSortOrderPrefix})\s*([a-zA-Z0-9]+)\s*(Notes|Comments|Formulas)(\.csv)?$");
    static Regex patGroup = new Regex(@$"({patSortOrderPrefix})\s*([a-zA-Z0-9]*)\s*(Group|Managed(By)?)(\.csv)?$");
    static Regex patSymbolName = new Regex(@$"({patSortOrderPrefix})\s*([a-zA-Z0-9]*)\s*SymbolName(\.csv)?$");
    static Regex patDecoratedSymbol = new Regex(@$"({patSortOrderPrefix})\s*([a-zA-Z0-9]*)\s*DecoratedSymbol(\.csv)?$");
    static Regex patAccount = new Regex(@$"({patSortOrderPrefix})\s*([a-zA-Z0-9]*)\s*(Account|AccntType)(\.csv)?$");
    static Regex patAtr21Day = new Regex(@$"({patSortOrderPrefix})\s*([a-zA-Z0-9]*)\s*(ATR21Day)(\.csv)?$");
    static Regex patNetLiqPercent = new Regex(@$"({patSortOrderPrefix})\s*([a-zA-Z0-9]*)\s*(NetLiqPercent|DATR21((Accnt|ManagedBy)(Group)?)?Percent)(-full)?(\.csv)?$");
    static Regex patShares = new Regex(@$"({patSortOrderPrefix}) ?([- a-zA-Z0-9]*)\s*[sS]hares(\.csv)?$");
    static Regex patUnrealizedGains = new Regex(@$"({patSortOrderPrefix}) ?([- a-zA-Z0-9]*)\s*(Daily ?Change|IRR|Price|DATR21|Shares|UnitCost|TotalCost|ProfitLoss(Dollar)?|NetLiquidValue|PortionOf(Total)?Account)(-full)?(\.csv)?$");
    static Regex patDailyChange = new Regex(@"Daily ?Change");
    static Regex patHistory = new Regex(@"--[Hh]istoryDays=([0-9]+)");
    static Regex patAddSpaceNotesSwitch = new Regex(@"^(.*)(Notes)$");
    static Regex patNoteColumnWidthSwitch = new Regex(@"--Note(Col(umn)?Width)?=([0-9]+)");
    static Regex patMaxEventAge = new Regex(@"--MaxEvent(Days?)?Age?=([0-9]+)");
    static Regex patOrderBySwitch = new Regex(@"--OrderBy=([a-zA-Z0-9_%\(\)]+)");
    static Regex patGenerateMasterSymbolList = new Regex(@"--([gG]enerate)?[mM]aster(Symbol)?List");
    static Regex patColorBySwitch = new Regex(@"--ColorBy=([a-zA-Z0-9_%/\(\)]+)");
    static Regex patMaxHistory = new Regex(@"--MaxHistory=([0-9]+)");
    static Regex patMaxNotesDaysOld = new Regex(@"--MaxNotesDaysOld=([0-9]+(\.[0-9]+))");
    static Regex patBackHistoryDayCount = new Regex(@"--BackHistoryDays?(Count)?=([0-9]+)");
    static Regex patTickerSymbolColumnWidth = new Regex(@"--(Ticker)?Symbol(Col(umn)?Width)?=([0-9]+)");
    static Regex patRemoveFinalM = new Regex(@"^(.*)M$");
    static Regex patRemoveLeadingZeros = new Regex(@"^(0+)(.*)$");
    static Regex patCheckStockFileLists = new Regex(@"--[Cc]heck(Stock(File(Lists)?)?)?");
    static readonly Regex patMasterStockList = new Regex("--[Mm]asterStockList(File)?$"); // beware! these are are weird: they don't have an "="
    static readonly Regex patMasterETFList = new Regex("--[Mm]asterETFList(File)?$"); // beware! these are are weird: they don't have an "="
    static readonly Regex patFileNameHasVersion = new Regex(@"^(.*)( \(([0-9]+)\))(\.csv)$");
    static Regex patOutputFileName = new Regex(@"--[Oo]ut(put(File(Name)?)?)?=([-A-Za-z0-9:\\%_\.]+)$");
    static Regex patBrokerageAccountEquitySymbol = new Regex(@"^([^:]+):([^:]+):([^:]+)$");
    static string NoteColumnWidth = "400";
    static string SymbolColumnWidth = "37";//"32.75";
    static List<Regex> patSkipFiles = new List<Regex> {
        //new Regex(@"sss Schwab NetLiqPercent(-(full|partial))?\.csv"),
        //new Regex(@"\([0-9]+\)\.csv$"),
        new Regex(@"^ETF Indicies\.csv$"),
        new Regex(@"Consolidated."),
        //new Regex(@"^sss Schwab (NetLiqPercent|DATR21AccntPercent|DATR21)(-(full|partial))\.csv$"),
        new Regex(@"^sss Schwab Shares(-(full|partial))\.csv$"),
        new Regex(@"ttt (SellSignalMCP Formulas|SellSignal200DayMA Comments)\.csv"),
        new Regex(@"^MinDollarVol10MComp50(ETF)?( \([0-9]+\))?\.csv$"),
        new Regex(@"^197 Industry Groups\.csv$"),
        new Regex(@"^All_Accounts_GainLoss_Realized(_Details)?_[0-9]+-[0-9]+.csv"),// Schwab TOS download
        new Regex(@"^Individual_[^_]+_Transactions_[0-9]+-[0-9]+.csv"),// Schwab web site download
        new Regex(@"Merrill-Holdings-Unrealized-Gain-Loss-Summary\.csv$"),
        new Regex(@"Merrill-Holdings\.csv$"),
        new Regex(@"^Merrill.*(Holdings?|All).*\.csv$"),
        new Regex(@"^ExportData.*\.csv$"),
        new Regex(@"(Account|Position)Statement\.csv$"),
        new Regex(@"^All_Accounts_GainLoss_Realized_Details_([0-9]+-[0-9]+)\.csv$"),
        new Regex(@"^Realized Gain Loss.*\.csv$"),
        new Regex(@"^UC_[0-9]+_[0-9]+\.csv$"), // Schwab TOS download
        new Regex(@"^[0-9]+-[0-9]+-[0-9]+-([0-9]+-)?(PositionStatement|watchlist)(-full)?\.csv$"), // Schwab TOS download
        new Regex(@"^Top (25 Funds Over [0-9]+ Years|Mutual Funds|[0-9]+ Industry or Sector Funds Over [0-9]+ Years)( \([0-9]+\))?\.csv"), // not used for comparison
        new Regex(@"^Top 25 Funds Over 10 years( \([0-9]+\))?\.csv"), // not used for comparison
        new Regex(@"^Top 25 Sector Funds Over 3 Years( \([0-9]+\))?\.csv$"), // not used for comparison
        new Regex(@"^Top Rated (Value|Growth) Funds( \([0-9]+\))?\.csv$"), // not used for comparison
        new Regex(@"^SEP-IRA-Positions-\d+-\d+-\d+-\d+\.csv$"),
        new Regex(@"^(197 Industry Groups|MarketSmith Growth 250)\.csv$"),
        //new Regex(@"^MarketSurge Growth 250\.csv$"),
        new Regex(@"2024 Aug Holdings Activity.csv$"),
        new Regex(@"[0-9]+-[0-9]+-[0-9]+-TradeActivity(-full)?\.csv$"), // schwab TOS download
        new Regex(@"^paper trading realized gain loss TOS \d+-\d+-\d+-AccountStatement\.csv$"),
        new Regex(@"Lot-Details\.csv$"),
        new Regex(@"^(Merril.*|SEP-IRA.*|paper trading.*|.*Ameritrade.*|197 Industry Groups|MinDollarVol[0-9]+MComp[0-9]+)\.csv$"),
    };
    static List<Regex> patSkipFilesMasterList = new List<Regex> { new Regex(@"^zzz.*\.csv$"), new Regex(@"^sss Schwab [A-Za-z0-9 ]+\.csv$") };
    static DateTime TODAY = System.DateTime.Today.AddHours(12);
    static bool SkipFile(string filePath)
    {
        var fileName = Path.GetFileName(filePath);
        foreach (var regex in patSkipFiles)
        {
            if (regex.IsMatch(fileName))
                return true;
        }
        return false;
    }
    static bool MatchSwitchWithValue(Regex pattern, string arg, out string switchValue, int captureGroup = 1)
    {
        var match = pattern.Match(arg);
        if (match.Success)
        {
            switchValue = match.Groups[captureGroup].Value;
            return true;
        }
        else
        {
            // fix this: should return null
            //switchValue = null; 
            switchValue = string.Empty;
            return false;
        }
    }
    static bool generateMasterRegexList = false;
    static double maxNotesDaysOld = 40;
    static int maxNoteSize = 1000;
    AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*ageMetric name*/, string/*ageMetric value*/>>> defaultTodayValues=new();
    AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*ageMetric name*/, string/*ageMetric value*/>>> defaultTodayETFValues=new();
    AutoMultiDimSortedDictionary<string/*file name*/, AutoMultiDimSortedDictionary<string /*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*ageMetric name*/, string/*ageMetric value*/>>>> mapFileNameToSymbolToDatesToAttributes = new();// new DescendingComparer<string>());
    ExcelStyleSaturationRange stockExcelSaturationAgeStyle;
    static AutoInitDoubleSortedDictionary<string> doubleMin = new();
    static AutoInitDoubleSortedDictionary<string> doubleMax = new();
    bool CheckStockFileLists {get; set; } = false;
    bool LoadExcelSpreadsheet = true;
    string? outputFileName = null;
    SortedDictionary<string, SymbolInList> masterSymbolListForWorksheet;

    // @@todo@@: remove thismapSortOrderPrefixedFileNameToColumnPosition. It is redundant
    SortedDictionary<string, int> mapSortOrderPrefixedFileNameToColumnPosition = new(new DescendingComparer<string>()); // key is file name only, value is column position in the worksheet. File name may be augmented with sort order prefix
    SortedDictionary<string, string> mapSortOrderPrefixedFileNameToFileName = new();
    SortedDictionary<string, string> mapFileNameToSortOrderPrefixedFileName = new();

    SortedDictionary<string, string> mapFileNameToColumnHeader = new();

    // @@todo@@ this is superfluous! Remove this and mapSortOrderPrefixedFileNameToColumnPosition
    string[] sortOrderPrefixedFileNames;

    bool fullPositionStatement = false; // true if the position statement contains all accounts
    static readonly Regex patFullOutputFileName = new Regex(@"-full\.xml$"); // if the output file name ends with -full.xml then the full position statement is being used.

    List<(string orderBy, string colorCodedAttributeName)> workSheets = new ();

    public static void Main(string[] args)
    {
        var main = new CompareTickerSymbolListsInCSVFilesMainProgram();
        int backHistoryDayCount = 250;
        int maxEventAge = 0;
        string initialRow = args.Length > 0 && Int32.TryParse(args[0], out Int32 _) ? args[0] : "6";
        var generateCSV = false;
        var generateXML = true;
        
        AutoMultiDimSortedDictionary<string/*symbol*/, AutoInitSortedDictionary<string/*file name only*/, SortedSet<DateTime>>> mapSymbolToFileNameToDates = new();
        SortedDictionary<string, DateTime> mapFileNameToMostRecentFileDate = new();
        SortedDictionary<DateTime, string> mapMostRecentDateToFile = new(new DescendingComparer<DateTime>());
        var sbStyles = new StringBuilder();
        var sbCSV = new StringBuilder();
        var symbolListForMasterRegex = new List<string>();
        var workSheetName = "Sort by ";
        var debug = false;
        var daysOfHistory = 5;
        var columnCurrent = 0;
        args = GetFileNamesAndSwitchesFromArgs(args);
        var verticalSplitter = false;
        var excelStyles = new AutoInitSortedDictionary<string/* Hexadecimal color name */, ExcelStyle>();
        // emacs solve([0 * m + b = 120, 1 * m + b = 235], [m, b])==[m = 115, b = 120]
        //var stockExcelHueAgeStyle = new ExcelStyleHueRange { Saturation = 240f, Luminance = 240f, HueMin = 120.0 / 360.0 * 255.0, HueMax = 235 / 360.0 * 255.0, InputMetricMin = 0, InputMetricMax = 1, HueScale = 115.0, HueOffset = 120 };
        main.stockExcelSaturationAgeStyle = new ExcelStyleSaturationRange { Hue = 120.0 / 360.0 * 255.0, Luminance = 240f, SaturationMin = 120.0 / 360.0 * 255.0, SaturationMax = 355.0 / 360.0 * 255.0, InputMetricMin = 0, InputMetricMax = 1, SaturationScale = 115.0, SaturationOffset = 120 };
        var mapAttributeNameToScreenTip = new Dictionary<string, (bool tip, string displayName, Func<string, string> convert)>();
        AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*ageMetric name*/, string/*ageMetric value*/>>>? emptyDefaultValues = null;

        var masterStockListFilePath = ExpandEnvironmentVariables(@"%DN%\MinDollarVol10MComp50.csv");
        var masterETFListFilePath = ExpandEnvironmentVariables(@"%DN%\MinDollarVol10MComp50ETF.csv");
        var csvDataFilesLoadList = new List<string>();

        // Throw away these minimum and maximum values because they are not from the lists were are interested in.
        doubleMin = new(); doubleMax = new();

        mapPositionToAttributeName.ToList().ForEach(e => mapAttributeNameToScreenTip.Add(e.Value.name, (e.Value.screenTip, e.Value.displayName, e.Value.convert)));

        IEnumerable<History> history = null;
        SortedSet<DateTime> historyDates = new(new DescendingComparer<DateTime>());
        SortedSet<string> fileNamesWithHistory = new();

        // this is never used
        //foreach((var position, (var name, var numeric)) in mapPositionToAttributeName)  mapAttributeNameToPosition.Add(name, position);
        SortedSet<string> requiredFiles = new();
        foreach (var pf in preferredFiles)
            if (!SkipFile(pf))
                requiredFiles.Add(pf);
            else if (!generateMasterRegexList)
                WriteLine($"skipping file {pf}");
        var idxArgs = 0;
        bool addSkipFiles = true;
        while (idxArgs < args.Length)
        {
            var arg = args[idxArgs];
            var matchHistory = patHistory.Match(arg);
            if (arg == "--debug")
            {
                debug = true;
            }
            else if (arg == "--skipConfidential")
            {
                skipConfidential = true;
            }
            else if (arg == "--help")
            {
                WriteLine("Help: TBD");
            }
            else if (patCheckStockFileLists.Match(arg).Success)
            {
                main.CheckStockFileLists = true;
            }
            else if (matchHistory.Success)
            {
                daysOfHistory = Int32.Parse(matchHistory.Groups[1].Value);
            }
            else if (arg == "--Reverse")
            {
                DescendingComparer<string>.reverse = true;
                verticalSplitter = true;
            }
            else if (arg == "--VerticalSplitter")
                verticalSplitter = true;
            else if (MatchSwitchWithValue(patOrderBySwitch, arg, out string orderByTemp, 1))
            {

                var colorCodedAttributeName = "";
                if (idxArgs + 1 < args.Length && MatchSwitchWithValue(patColorBySwitch, args[idxArgs + 1], out string colorByTemp, 1))
                {
                    colorCodedAttributeName = colorByTemp;
                    idxArgs++;
                    arg = args[idxArgs];
                }
                main.workSheets.Add((orderBy: orderByTemp.Replace('_', ' '), colorCodedAttributeName: colorCodedAttributeName.Replace('_', ' ')));
                SymbolColumnWidth = "55";
            }
            else if (arg == "--noWildCards")
            {
                // skipBaseFile, already handled
            }

            else if (patGenerateMasterSymbolList.Match(arg).Success /*fileName == "--masterList"*/)
            {
                generateMasterRegexList = true;
                patSkipFiles.Add(new Regex(@"sss Schwab NetLiqPercent(-(full|partial))?\.csv")); // this has equity symbols that include Siegfried's special prefixes for account and group (separated with colons) and this will mess up the generation of a master symbol list
                patSkipFiles.AddRange(patSkipFilesMasterList);
                generateXML = false;
                generateCSV = false;
            }
            else if (MatchSwitchWithValue(patNoteColumnWidthSwitch, arg, out string width, 3))
            {
                NoteColumnWidth = width;
            }
            else if (MatchSwitchWithValue(patMaxEventAge, arg, out string tmpMaxEventAge, 2))
            {
                maxEventAge = int.Parse(tmpMaxEventAge);
            }
            else if (MatchSwitchWithValue(patGotoRow, arg, out string row))
            {
                initialRow = row;
            }
            else if (MatchSwitchWithValue(patMaxNotesDaysOld, arg, out string daysOld))
            {
                maxNotesDaysOld = double.Parse(daysOld);
            }
            else if (MatchSwitchWithValue(patBackHistoryDayCount, arg, out string tmpBackHistoryDayCount, 2))
            {
                backHistoryDayCount = Int32.Parse(tmpBackHistoryDayCount);
            }
            // else if (MatchSwitchWithValue(patOutputFileName, fileName, out string fileName, 4))
            // {
            //     if (!string.IsNullOrEmpty(fileName))
            //         main.outputFileName = fileName;
            // }
            // I could not make this work when there were spaces in the file name.
            //else if (MatchSwitchWithValue(patMasterStockList, fileName, out string masterStockListFileName, 2))
            //{
            //    if (!string.IsNullOrEmpty(masterStockListFileName))
            //    {
            //        masterStockListFilePath = ExpandEnvironmentVariables(masterStockListFileName);
            //        if (!generateMasterRegexList)
            //            WriteLine($"masterStockListFilePath = {masterStockListFilePath}");
            //    }
            //}
            //else if (MatchSwitchWithValue(patMasterETFList, fileName, out string masterETFListFileName, 2))
            //{
            //    if (!string.IsNullOrEmpty(masterETFListFileName))
            //    {
            //        masterETFListFilePath = ExpandEnvironmentVariables(masterETFListFileName);
            //        if(!generateMasterRegexList)
            //            WriteLine($"masterETFListFilePath = {masterETFListFilePath}");
            //    }
            //}
            else if (arg == "--outputFileName")
            {
                var fn = args[++idxArgs];
                if (!string.IsNullOrEmpty(fn))
                {
                    main.outputFileName = ExpandEnvironmentVariables(fn);
                    main.fullPositionStatement = patFullOutputFileName.Match(Path.GetFileName(main.outputFileName)).Success;
                    if (!generateMasterRegexList)
                        WriteLine($"main.outputFileName = {main.outputFileName}");
                }
            }
            else if (patMasterStockList.Match(arg).Success /*arg == "--masterStockList"*/) // watch out! There is no "=" and this is inconsistent with the other switches.
            {
                var fn = args[++idxArgs];
                if (!string.IsNullOrEmpty(fn))
                {
                    masterStockListFilePath = ExpandEnvironmentVariables(fn);
                    //ConditionallyAddUnVersionedFileNameToSkipList(fn);
                    if (!generateMasterRegexList)
                        WriteLine($"masterStockListFilePath = {masterStockListFilePath}");
                }
            }
            else if (patMasterETFList.Match(arg).Success/*arg == "--masterETFList"*/) // watch out! There is no "=" and this is inconsistent with the other switches.
            {
                var fn = args[++idxArgs];
                if (!string.IsNullOrEmpty(fn))
                {
                    masterETFListFilePath = ExpandEnvironmentVariables(fn);
                    var m = patFileNameHasVersion.Match(Path.GetFileName(fn));
                    //ConditionallyAddUnVersionedFileNameToSkipList(fn);
                    if (!generateMasterRegexList)
                        WriteLine($"masterETFListFilePath = {masterETFListFilePath}");
                }
            }
            else if (MatchSwitchWithValue(patTickerSymbolColumnWidth, arg, out string tickerSymbolColumnWidth, 4))
            {
                SymbolColumnWidth = tickerSymbolColumnWidth;
            }
            else // it is not a command line switch. It must be a file name
            {
                if (addSkipFiles)
                {
                    // assumption: we have processed the --outputFileName switch if it was present and determined if the full position statement is being used or not.
                    // Assumption: the full position statement file name contains all accounts and the partial position statement file name does not contain all of them.
                    // @@todo@@: --outputFileName is now required if we are going to deal with partial v full position statements.
                    // what to do? error out if --outputFileName is not present?
                    var patstr = @"^(sss Schwab )?(Net\s*Liq|DATR21)\s*(Net\s*Liq|(Accnt|ManagedBy)(\s*Group)?)?(\s*Percent)?" + (main.fullPositionStatement ? "" : "-full") + @"\.csv$";
                    var pat = new Regex(patstr);
                    //var xyz = pat.Match("sss Schwab DATR21-full.csv").Success;
                    patSkipFiles.Add(pat);
                    addSkipFiles = false;
                }


                var fn = Path.GetFileName(arg);
                (string baseNameNoVersionFn, bool versioned) = GetBaseFileNameNoVersions(fn);
                var annotation = versioned ? $"({Path.GetFileName(baseNameNoVersionFn)})" : "";

                bool skipBaseFile = SkipFile(baseNameNoVersionFn);
                bool skipFile = SkipFile(fn);
                if (versioned && requiredFiles.Contains(baseNameNoVersionFn))
                {
                    requiredFiles.Remove(baseNameNoVersionFn); // Keep a running list of remaining files to be loaded
                    if (!generateMasterRegexList && !main.CheckStockFileLists)
                        WriteLine($"Found required file '{baseNameNoVersionFn}'");
                }
                else if (!versioned && requiredFiles.Contains(fn))
                {
                    requiredFiles.Remove(fn); // Keep a running list of remaining files to be loaded
                    if (!generateMasterRegexList && !main.CheckStockFileLists)
                        WriteLine($"Found required file '{fn}'");
                }
                else if (!versioned & !generateMasterRegexList && !main.CheckStockFileLists && !skipFile)
                {
                    WriteLine($"Found extra file '{fn}'");
                }
                else if (versioned && !generateMasterRegexList && !main.CheckStockFileLists && !skipBaseFile)
                {
                    WriteLine($"Found extra file '{baseNameNoVersionFn}'");
                }
                if (!main.CheckStockFileLists)
                {
                    if (skipConfidential && confidential.Contains(baseNameNoVersionFn) && !generateMasterRegexList)
                    {
                        WriteLine($"Skipping confidential file '{baseNameNoVersionFn}'");
                    }
                    else if (skipBaseFile)
                    {
                        if (!generateMasterRegexList)
                            WriteLine($"Skipping file '{baseNameNoVersionFn}'");
                    }
                    else if (skipFile)
                    {
                        if (!generateMasterRegexList)
                            WriteLine($"Skipping file '{fn}'");
                    }
                    else
                    {
                        csvDataFilesLoadList.Add(arg);
                    }
                }
            }
            idxArgs++;
        }

        RemoveOlderVersionedFileNamesFromFilesList(csvDataFilesLoadList);

        bool flowControl = true; // ExitIfDownloadsDirectoryContainsParens();
        if (!flowControl && !generateMasterRegexList)
        {
            WriteLine("Downloads directory files contain parentheses. Please remove them and try again.");
            // exist with failure code
            Environment.Exit(1);
        }

        // add the only the file name without the path and extension to the preferred files list
        foreach (var file in new[] { masterStockListFilePath, masterETFListFilePath })
        {
            if (File.Exists(file))
            {
                var fileName = Path.GetFileName(file);
                if (!string.IsNullOrEmpty(fileName) && !preferredFiles.Contains(fileName))
                {
                    preferredFiles.Add(fileName);
                }
            }
            //if (!SkipFile(file)) requiredFiles.Add(file);
            patSkipFiles.Add(new Regex(Regex.Escape(file)
                //, RegexOptions.IgnoreCase
                )); // the master files are special csv files that are not to be processed like the other csv files.
        }

        try
        {
            main.MergeParseAndLoadMasterLists(maxEventAge, emptyDefaultValues, masterStockListFilePath, masterETFListFilePath);
        }
        catch (FileNotFoundException ex)
        {
            WriteLine($"Missing required file: {ex.FileName}");
            return;
        }

        csvDataFilesLoadList.ForEach(fn =>
            main.LoadCSVDataFiles(backHistoryDayCount, maxEventAge, generateMasterRegexList, /*main.mapSortOrderPrefixedFileNameToColumnPosition,*/ mapSymbolToFileNameToDates, mapFileNameToMostRecentFileDate, mapMostRecentDateToFile, ref columnCurrent, emptyDefaultValues, ref history, ref historyDates, fileNamesWithHistory, fn));

        if (requiredFiles.Count > 0 && !generateMasterRegexList)
        {
            WriteLine($"{requiredFiles.Count} Required files not found: {string.Join(", ", requiredFiles)}");
            if (main.CheckStockFileLists)
                return;
        }
        else if (main.CheckStockFileLists && !generateMasterRegexList)
        {
            WriteLine("All required files found");
            return;
        }

        /*SortedDictionary<string, SymbolInList>*/
        main.masterSymbolListForWorksheet = main.MakeMasterSymbolListForWorksheet(mapFileNameToMostRecentFileDate, debug);
        if (generateMasterRegexList) symbolListForMasterRegex = MakeSymbolListForMasterRegex(main.masterSymbolListForWorksheet);
        if (generateMasterRegexList && symbolListForMasterRegex.Count() > 0)
        {
            WriteLine(string.Join("|", symbolListForMasterRegex)
                + "|" + symbolListForMasterRegex.Last() // bug fix for perl
                );
        }
        var stStyles = "";
        main.sortOrderPrefixedFileNames = MakeFileNamesArray(main.mapSortOrderPrefixedFileNameToColumnPosition); // main.mapSortOrderPrefixedFileNameToColumnPosition (keys only) populated in LoadCSVDataFiles
        main.mapFileNameToColumnHeader = MakeMapFileNameToColumnHeader(main.sortOrderPrefixedFileNames);
        var finalExcelOutputFilePath = string.IsNullOrEmpty(main.outputFileName) ? ExpandEnvironmentVariables(@$"%USERPROFILE%\Downloads\CompareMarketSmithLists_{DateTime.Now.ToString("yyyy-MMM-dd-ddd-HH")}.{(generateCSV || generateMasterRegexList ? "csv" : "xml")}") : ExpandEnvironmentVariables(main.outputFileName);
        if (!generateMasterRegexList)
            WriteLine($"finalOutputFilename = {finalExcelOutputFilePath}");
        var saveColumnCurrent = columnCurrent;
        List<string> workSheetXML = main.MakeLegendWorksheet();
        var worksheetCount = 0;
        foreach (var workSheet in main.workSheets)
        {
            var sbXMLWorksheetRows = new StringBuilder();
            var xmlColumnWidths = new List<string> { """<Column ss:Width="15" />""" }; // why does the width have no affect?
            int columnCount = main.GenerateXMLWorksheetColumnDeclarations(/*main.mapSortOrderPrefixedFileNameToColumnPosition,*/ sbXMLWorksheetRows, sbCSV, xmlColumnWidths/*, main.sortOrderPrefixedFileNames*/);
            var historyDirectoryDateArray = historyDates.ToArray<DateTime>().Reverse<DateTime>().ToArray<DateTime>(); // why do I have to reverse this since it was created with a descending comparitor?
            var rowCount = 0;
            columnCurrent = saveColumnCurrent;
            main.GenerateXMLWorksheetDataRows(backHistoryDayCount,
                //main.mapFileNameToSymbolToDatesToAttributes, 
                main.masterSymbolListForWorksheet,
                //mapSortOrderPrefixedFileNameToColumnPosition,
                //mapFileNameToColumnHeader, 
                mapSymbolToFileNameToDates, sbXMLWorksheetRows, sbCSV, debug, ref rowCount, ref columnCurrent, excelStyles, fileNamesWithHistory, columnCount, historyDirectoryDateArray, mapAttributeNameToScreenTip, workSheet);
            var columnHeaders = $"""<Cell ss:StyleID="s62"><Data ss:Type="String">Order</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>"""
                + $"""{string.Join("", (main.sortOrderPrefixedFileNames.OrderBy(fn => main.mapSortOrderPrefixedFileNameToColumnPosition[fn])).Select(fn => $"""<Cell ss:StyleID="s62"><Data ss:Type="String">{patAddSpaceNotesSwitch.Replace(main.mapFileNameToColumnHeader/*mapFileNameToColumnHeader*/[fn], m => $"{m.Groups[1].Value} {m.Groups[2].Value}")}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>"""))} """
                + "";
            columnHeaders += $"""{string.Join("", (mapPositionToAttributeName.Select(positionNamePair => $"""<Cell ss:StyleID="s62"><Data ss:Type="String">{positionNamePair.Value.displayName}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>""")))}""";
            columnCount += mapPositionToAttributeName.Count;
            rowCount++;
            columnCount++; // include the extra column for the current row
            string workSheetDisplayName = "" + (++worksheetCount) + ". " + workSheetName + workSheet.orderBy.Trim() + (string.IsNullOrEmpty(workSheet.colorCodedAttributeName) ? "" : "-ColorBy" + workSheet.colorCodedAttributeName.Replace("/", ""));
            if (workSheetDisplayName.Length > 31) workSheetDisplayName = workSheetDisplayName.Substring(0, 31); // Excel worksheet names are limited to 31 characters
            var len = workSheetDisplayName.Length;
            workSheetXML.Add(ComposeWorkSheet(initialRow, sbXMLWorksheetRows.ToString(), workSheetDisplayName, rowCount, xmlColumnWidths, verticalSplitter, columnCount, columnHeaders));
        }
        if (!generateMasterRegexList)
        {
            foreach ((var _, var v) in excelStyles) stStyles += (string.IsNullOrEmpty(stStyles) ? "" : "\n    ") + v;
            var workBookXML = $"""
             <?xml version="1.0"?>
             <?mso-application progid="Excel.Sheet"?>
             <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40">
             <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
              <Author>Siegfried Heintze</Author>
              <LastAuthor>Siegfried Heintze</LastAuthor>
              <Created>2022-09-27T16:45:28Z</Created>
              <Version>16.00</Version>
             </DocumentProperties>
             <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
              <AllowPNG/>
             </OfficeDocumentSettings>
             <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
              <WindowHeight>28380</WindowHeight>
              <WindowWidth>32767</WindowWidth>
              <WindowTopX>32767</WindowTopX>
              <WindowTopY>32767</WindowTopY>
              <ProtectStructure>False</ProtectStructure>
              <ProtectWindows>False</ProtectWindows>
             </ExcelWorkbook>
             <Styles>
              <Style ss:ID="Default" ss:Name="Normal">
               <Alignment ss:Vertical="Center" ss:WrapText="1"/>
               <Borders/>
               <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
               <Interior/>
               <NumberFormat/>
               <Protection/>
              </Style>
              <Style ss:ID="s18" ss:Name="Currency">
                <NumberFormat ss:Format="_(&quot;$&quot;* #,##0.00_);_(&quot;$&quot;* \(#,##0.00\);_(&quot;$&quot;* &quot;-&quot;??_);_(@_)"/>
              </Style>
              <Style ss:ID="s20" ss:Name="Percent">
                <NumberFormat ss:Format="0%"/>
              </Style>
              <Style ss:ID="s62">
               <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:Rotate="90" ss:WrapText="1"/>
              </Style>
              <Style ss:ID="s63">
               <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
              </Style>
              <Style ss:ID="s65">
               <NumberFormat ss:Format="Fixed"/>
              </Style>
              <Style ss:ID="s66">
                <NumberFormat ss:Format="ddd\ d\-mmm\-yy"/>
              </Style> 
              <Style ss:ID="s67">                
              </Style>
              <Style ss:ID="s68" ss:Name="Hyperlink">
                <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#0066CC" ss:Underline="Single"/>
              </Style>
              <Style ss:ID="s69">
                <NumberFormat/>
              </Style>
              <Style ss:ID="s70" ss:Parent="s20">
                <Alignment ss:Vertical="Center" ss:WrapText="1"/>
                <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
                <NumberFormat ss:Format="Percent"/>
              </Style>
              <Style ss:ID="s71" ss:Parent="s20">
                <Alignment ss:Vertical="Center" ss:WrapText="1"/>
                <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#FF0000"/>
                <NumberFormat ss:Format="Percent"/>
              </Style>
              <Style ss:ID="s72" ss:Parent="s18">
               <Alignment ss:Vertical="Center" ss:WrapText="1"/>
               <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
              </Style>
              <Style ss:ID="s73" ss:Parent="s20">
                <Alignment ss:Vertical="Center" ss:WrapText="1"/>
                <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#FF0000" ss:Bold="1"/>
                <NumberFormat ss:Format="Percent"/>
              </Style>
              <Style ss:ID="s74" ss:Parent="s18">
               <Alignment ss:Vertical="Center" ss:WrapText="1"/>
               <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#FF0000"/>
              </Style>
              <Style ss:ID="s75Percent" ss:Parent="s68">
                <NumberFormat ss:Format="Percent"/>
              </Style>
              <Style ss:ID="s76Comma" ss:Name="Comma">
                <NumberFormat ss:Format="_(* #,##0.00_);_(* \(#,##0.00\);_(* &quot;-&quot;??_);_(@_)"/>
              </Style>
              {stStyles}
             </Styles>

             """ +
         string.Join("", workSheetXML) +
         """"
            </Workbook>            
            """";

            //WriteLine(worksheetXML);
            var done = false;
            while (!done)
            {
                try
                {
                    System.IO.File.WriteAllText(finalExcelOutputFilePath, workBookXML);
                    done = true;
                }
                catch (System.IO.IOException ex)
                {
                    finalExcelOutputFilePath = ExpandEnvironmentVariables(@$"%USERPROFILE%\Downloads\CompareMarketSmithLists_{DateTime.Now.ToString("yyyy-MMM-dd-ddd-HH-mm-ss")}.{(generateCSV || generateMasterRegexList ? "csv" : "xml")}");
                    WriteLine($"Error writing file {ex.Message}, trying again with new file name {finalExcelOutputFilePath} sleeping...");
                    Thread.Sleep(1000);
                }
                catch (Exception ex)
                {
                    WriteLine($"{ex}");
                    done = true;
                }
            }
            //WriteLine($"files processed={string.Join(",\n", mapSortOrderPrefixedFileNameToColumnPosition.Keys.Select(fileName=>$"@\"{fileName}\"").ToList())}");

            //var excel = ExpandEnvironmentVariables(@"%MSOFFICE%\EXCEL.EXE");
            //Process.Start(excel, $"/s \"{finalExcelOutputFilePath}\"");


            // add Microsoft.Office.Interop.Excel.Application to this project
            // add Microsoft.Office.Interop.Excel to the project references

            // Excel COM Interop is missing DLLs. 
            // https://stackoverflow.com/questions/32399420/could-not-load-file-or-assembly-office-version-15-0-0-0
            /*
            // start Microsoft Excel as an out-of-process COM object
            var excel = new Microsoft.Office.Interop.Excel.Application();
            // make Excel visible
            excel.Visible = true;
            var workbook = excel.Workbooks.Open(finalExcelOutputFilePath);
            //SortMultipleWorksheets(workbook, excel);
            // save as xlsx
            workbook.SaveAs(finalExcelOutputFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
            // close the workbook
            workbook.Close();
            */

            // exit with success code
            Environment.Exit(0);


        }
    }

    static (string,bool) GetBaseFileNameNoVersions(string fn)
    {
        var m = patFileNameHasVersion.Match(fn);
        var baseNameNoVersionFn = m.Success ? m.Groups[1].Value + m.Groups[4].Value : fn;
        return (baseNameNoVersionFn, m.Success);
    }

    /// <summary>
    /// When two file names are the same except for a version number (as per the browser convenction of appending an integer enclosed by parentheses to the file name before the extension) delete the older one from the list.
    /// </summary>
    /// <returns></returns>
    static void RemoveOlderVersionedFileNamesFromFilesList(List<string> dataFiles)
    {
        var mapBaseFileNameToFullFileNameAndVersion = new Dictionary<string, (string fullFileName, int version)>();
        foreach (var filePath in dataFiles)
        {
            var fileName = Path.GetFileName(filePath);
            var m = patFileNameHasVersion.Match(fileName);
            if (m.Success)
            {
                var baseFileName = m.Groups[1].Value + m.Groups[4].Value;
                var version = int.Parse(m.Groups[3].Value);
                if (mapBaseFileNameToFullFileNameAndVersion.ContainsKey(baseFileName))
                {
                    var (existingFullFileName, existingVersion) = mapBaseFileNameToFullFileNameAndVersion[baseFileName];
                    if (version > existingVersion)
                    {
                        if(!generateMasterRegexList)
                            WriteLine($"remove older versioned file {existingFullFileName} in favor of newer versioned file {filePath}");
                        // replace with the newer version
                        mapBaseFileNameToFullFileNameAndVersion[baseFileName] = (filePath, version);
                    }
                }
                else
                {
                    mapBaseFileNameToFullFileNameAndVersion[baseFileName] = (filePath, version);
                }
            }
            else
            {
                // not a versioned file name, but there might already be a versioned file name for this base file name
                if (!mapBaseFileNameToFullFileNameAndVersion.ContainsKey(fileName))
                    mapBaseFileNameToFullFileNameAndVersion[fileName] = (filePath, 0);
                else
                {
                    //patSkipFiles.Add(new Regex("^" + Regex.Escape(fileName) + "$")); // skipBaseFile the unversioned file if the versioned file is present
                    if (!generateMasterRegexList)
                        WriteLine($"skipBaseFile file {fileName} because it superseded by {mapBaseFileNameToFullFileNameAndVersion[fileName].fullFileName}");
                }
                
            }
        }
        // now rebuild the dataFiles list
        dataFiles.Clear();
        foreach (var (fullFileName, _) in mapBaseFileNameToFullFileNameAndVersion.Values)
        {
            dataFiles.Add(fullFileName);
        }

    }


    List<string> MakeLegendWorksheet()
    {
        var sbLegend = new StringBuilder();
        var workSheetXML = new List<string>();
        sbLegend.Append("""<Row><Cell><Data ss:Type="String">Workseets</Data></Cell><Cell><Data ss:Type="String">Description</Data></Cell></Row>\n""");
        foreach (var worksheet in workSheets)
        {
            sbLegend.Append("""<Row><Cell><Data ss:Type="String">""");
            sbLegend.Append($"Sort by {worksheet.orderBy} {(string.IsNullOrEmpty(worksheet.colorCodedAttributeName) ? "" : $"-Color by {worksheet.colorCodedAttributeName}")}");
            sbLegend.Append("</Data></Cell></Row>\n");
        }
        sbLegend.Append("""<Row><Cell><Data ss:Type="String">Flags</Data></Cell><Cell><Data ss:Type="String">Meaning</Data></Cell></Row>\n""");
        foreach (var (key, value) in flags)
        {
            sbLegend.Append("""<Row><Cell><Data ss:Type="String">""");
            sbLegend.Append($"{key}");
            sbLegend.Append($"""</Data></Cell><Cell><Data ss:Type="String"> {value}  </Data> </Cell></Row>\n""");
        }
        var legend = $"""
            <Worksheet ss:Name="Legend">
             <Table ss:ExpandedColumnCount="2" ss:ExpandedRowCount="{workSheets.Count() + 1 + flags.Count() + 1}" x:FullColumns="1" x:FullRows="1" ss:DefaultRowHeight="15">
              <Column ss:AutoFitWidth="0" ss:Width="417"/>
              <Column ss:AutoFitWidth="0" ss:Width="417"/>
              {sbLegend.ToString()}
             </Table>
             <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
              <PageSetup>
               <Header x:Margin="0.3"/>
               <Footer x:Margin="0.3"/>
               <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
              </PageSetup>
              <Selected/>
              <Panes>
               <Pane>
                <Number>3</Number>
                <ActiveRow>9</ActiveRow>
               </Pane>
              </Panes>
              <ProtectObjects>False</ProtectObjects>
              <ProtectScenarios>False</ProtectScenarios>
             </WorksheetOptions>
            </Worksheet>
            """;
        workSheetXML.Add(legend);
        return workSheetXML;
    }

    private static void ConditionallyAddUnVersionedFileNameToSkipList(string fn)
    {
        var m = patFileNameHasVersion.Match(Path.GetFileName(fn));
        if (m.Success)
        {
            var noVersionFileName = m.Groups[1].Value + m.Groups[3].Value;
            patSkipFiles.Add(new Regex("^" + Regex.Escape(noVersionFileName) + "$")); // skipBaseFile the unversioned file if the versioned file is present
            if (!generateMasterRegexList)
                WriteLine($"adding unversioned file {noVersionFileName} to the skipBaseFile list");
        }
    }

    static bool ExitIfDownloadsDirectoryContainsParens()
    {
        // Enumerate all the files in the Downloads directory and exit if any of the file names contain parentheses.
        // This is to prevent the user from running this program with a file name that contains parentheses.
        var downloadsDirectory = ExpandEnvironmentVariables(@"%USERPROFILE%\Downloads");
        if (Directory.Exists(downloadsDirectory))
        {
            var files = Directory.EnumerateFiles(downloadsDirectory, "*.csv", SearchOption.TopDirectoryOnly)
                //.Where(fileName => !SkipFile(Path.GetFileName(fileName)))
                ;
            foreach (var file in files)
            {
                var fileName = Path.GetFileName(file);
                if(fileName.Contains("MinDollarVol10MComp50ETF") // special master list
                   || fileName.Contains("MinDollarVol10MComp50") // special master list
                   || fileName.Contains("Top Rated Growth Funds") // skipped
                   || fileName.Contains("Top Rated Value Funds")  // skipped
                   || fileName.Contains("Top 25 Sector Funds Over 3 Years") // skipped @@todo@@ check the skipBaseFile list instead of hard coding these!
                   || fileName.Contains("Top 25 Funds Over 10 years") // skipped
                   || fileName.Contains("ETF Indicies"))          // used (maybe) but we ignore the parens
                {
                    // skipBaseFile these files: we have special logic to handle these files when they include parentheses
                    continue;
                }
                if (fileName.Contains('(') || fileName.Contains(')'))
                {
                    WriteLine($"Error: File name '{fileName}' contains parentheses. Please rename the file and try again.");
                    return false;
                }
            }
        }

        return true;
    }

    /// <summary>
    /// Merges master lists of values from two CSV files based on a specified event age and optional default values. This is necessary because of a constraint with MarketSurge
    /// </summary>
    /// <param name="maxEventAge">Specifies the maximum age of events to consider when merging data.</param>
    /// <param name="emptyDefaultValues">Provides a structure for default values that can be used if no data is found.</param>
    /// <param name="defaultValuesFileName">Indicates the file name of the CSV containing default values to be parsed.</param>
    /// <param name="defaultETFValuesFileName">Indicates the file name of the CSV containing ETF default values to be parsed.</param>
    void MergeParseAndLoadMasterLists(int maxEventAge, AutoMultiDimSortedDictionary<string, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string, string>>>? emptyDefaultValues, string defaultValuesFileName,
        string defaultETFValuesFileName)
    {
        if(!generateMasterRegexList && !CheckStockFileLists)
            WriteLine($"Merging master lists from {defaultValuesFileName} and {defaultETFValuesFileName} with maxEventAge={maxEventAge}");
        defaultTodayValues = ParseCSV(defaultValuesFileName, ExtractDateTimeFromFilePath(defaultValuesFileName).Date, emptyDefaultValues, maxEventAge);
        defaultTodayETFValues = ParseCSV(defaultETFValuesFileName, ExtractDateTimeFromFilePath(defaultValuesFileName).Date, emptyDefaultValues, maxEventAge);
        foreach (var (symbol, mapDatesToMetrics) in defaultTodayETFValues)
            foreach (var (date, mapMetrics) in mapDatesToMetrics)
                foreach (var (metric, value) in mapMetrics)
                    defaultTodayValues[symbol][date][metric] = value;
        if (!generateMasterRegexList && !CheckStockFileLists)
            WriteLine($"Complete: Loaded {defaultTodayValues.Count} symbols with {defaultTodayValues.Values.SelectMany(d => d.Keys).Count()} dates from {defaultValuesFileName} and {defaultETFValuesFileName}");
    }

    void LoadCSVDataFiles(int backHistoryDayCount, int maxEventAge, bool generateMasterRegexList/*, SortedDictionary<string, int> mapSortOrderPrefixedFileNameToColumnPosition*/, AutoMultiDimSortedDictionary<string, AutoInitSortedDictionary<string, SortedSet<DateTime>>> mapSymbolToFileNameToDates, SortedDictionary<string, DateTime> mapFileNameToMostRecentFileDate, SortedDictionary<DateTime, string> mapMostRecentDateToFile, ref int columnCurrent, AutoMultiDimSortedDictionary<string, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string, string>>>? emptyDefaultValues, ref IEnumerable<History> history, ref SortedSet<DateTime> historyDates, SortedSet<string> fileNamesWithHistory, string fileName)
    {
        //var sortOrderPrefixedFileNames = ExpandFilePaths(fileName);
        fileName = ExpandEnvironmentVariables(fileName); // this fileName may contain the browser version number in parentheses        
        (var baseNameNoVersionFn, var versioned) = GetBaseFileNameNoVersions(/*Path.GetFileName*/(fileName));
        var annotation = versioned ? $"({Path.GetFileName(baseNameNoVersionFn)}) " : "";
        var skip = SkipFile(versioned ? baseNameNoVersionFn: fileName); // never process a file more than once

        if (File.Exists(fileName) && !skip) // @@todo@@ reverse order of conditions
        {
            try
            {
                //DateTime lastWriteTime = System.IO.File.GetLastWriteTime(fileName);
                var listDateTime = ExtractDateTimeFromFilePath(fileName);
                AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*ageMetric name*/, string/*ageMetric value*/>>> csvCurrentData = ParseCSV(fileName, listDateTime, defaultTodayValues, maxEventAge);
                if (!generateMasterRegexList)
                {
                    if(csvCurrentData.Count == 0)
                    {
                        WriteLine($"No data found in file {fileName}{annotation}, skipping");
                        return;
                    }
                    WriteLine($"Loaded file {fileName}{annotation} with {csvCurrentData.Count} symbols and {csvCurrentData.Values.SelectMany(d => d.Keys).Count()} dates");
                }                
                fileName = Path.GetFileName(fileName);
                baseNameNoVersionFn = Path.GetFileName(baseNameNoVersionFn);
                if (generateMasterRegexList)
                    history = new List<History>();
                else
                {
                    var downloadDirs = Directory.GetDirectories(ExpandEnvironmentVariables(@"%USERPROFILE%\Downloads")).ToList();
                    var downloadDirs2 = Directory.GetDirectories(ExpandEnvironmentVariables(@"%DN002%"));
                    downloadDirs.AddRange(downloadDirs2);
                    // assumption: there are no file names in the history that have browser version numbers
                    history = downloadDirs.Select(d => new { Name = Path.GetFileName(d), Path = d }).Where(
                        d => patDateTime_YYYY_MMM_dd_ddd.Match(d.Name).Success).Select(
                        d => { var m = patDateTime_YYYY_MMM_dd_ddd.Match(d.Name); var r = new { Name = d.Name, Path = d.Path, Date = System.DateTime.Parse(m.Groups[1].Value) }; return r; }).OrderByDescending(
                        d => d.Date).Take(backHistoryDayCount).Select(
                        f => new History { Date = f.Date, FileExists = System.IO.File.Exists(Path.Combine(f.Path + "\\" + fileName)), Name = System.IO.Path.Combine(f.Path + "\\" + fileName) });
                }
                mapMostRecentDateToFile[listDateTime.Date] = fileName;
                mapFileNameToMostRecentFileDate[fileName] = listDateTime.Date;
                MakeHistoryDateSet(history, TODAY, ref historyDates);
                UpdateMapSymbolToFileNameToDates(mapSymbolToFileNameToDates, baseNameNoVersionFn, csvCurrentData);
                UpdateMapFileNameToSymbolToDateToAttributes(mapFileNameToSymbolToDatesToAttributes, fileName, csvCurrentData);
                if (!generateMasterRegexList)
                {
                    foreach (var oldFile in history)
                    {
                        if (oldFile.FileExists) // this should not have version numbers
                        {
                            var csvHistoryData = ParseCSV(oldFile.Name, oldFile.Date, emptyDefaultValues, int.MaxValue);
                            foreach (var sym in csvCurrentData.Keys)
                            {
                                if (csvHistoryData.Keys.Contains(sym))
                                {
                                    csvCurrentData[sym][oldFile.Date] = csvHistoryData[sym][oldFile.Date];
                                }
                                else
                                {
                                    break;
                                }
                            }
                            var fn = Path.GetFileName(oldFile.Name);
                            UpdateMapSymbolToFileNameToDates(mapSymbolToFileNameToDates, fn, csvHistoryData);
                            UpdateMapFileNameToSymbolToDateToAttributes(mapFileNameToSymbolToDatesToAttributes, fn, csvHistoryData);
                            fileNamesWithHistory.Add(fn);
                        }
                    }
                }                
                var sortOrderPrefixedFileName = MaybeAddSortOrderPrefix(baseNameNoVersionFn);
                mapSortOrderPrefixedFileNameToColumnPosition[sortOrderPrefixedFileName] = 0; // Populate values later
                mapSortOrderPrefixedFileNameToFileName[sortOrderPrefixedFileName] = baseNameNoVersionFn;
                mapFileNameToSortOrderPrefixedFileName[fileName] = sortOrderPrefixedFileName;
				
                columnCurrent++;
            }
            catch (IOException ex)
            {
                WriteLine($"skipping file {fileName}{annotation} Error={ex}");
            }
        }
        else if (!generateMasterRegexList)
        {
            WriteLine($"{((skip ? "Skip file" : "File not found") + ": ")} {fileName}{annotation}");
        }
    }
    static string ComposeWorkSheet(string initialRow, string XMLWorksheetRows, string workSheetName, int rowCount, List<string> xmlColumnWidths, bool verticalSplitter, int columnCount, string columnHeaders)
    { // height was 230.25
        return $"""
             <Worksheet ss:Name="{workSheetName}">
               <Names>            
                <NamedRange ss:Name="_FilterDatabase" ss:RefersTo="='{workSheetName}'!R1C1:R{rowCount}C{columnCount}" ss:Hidden="1"/>
               </Names>
               <Table ss:ExpandedColumnCount="{columnCount}" ss:ExpandedRowCount="{rowCount}" x:FullColumns="1" x:FullRows="1" ss:DefaultColumnWidth="42" ss:DefaultRowHeight="11.25">
                 {string.Join("", xmlColumnWidths)}
                 <Row ss:Height="62.25">
                   {columnHeaders}
                 </Row>
                 {XMLWorksheetRows}
               </Table>
               <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
                <PageSetup>
                 <Header x:Margin="0.3"/>
                 <Footer x:Margin="0.3"/>
                 <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
                </PageSetup>
               <Unsynced/>
               <Print>
                <ValidPrinterInfo/>
                <HorizontalResolution>600</HorizontalResolution>
                <VerticalResolution>600</VerticalResolution>
               </Print>
               <Selected/>
             """ + (verticalSplitter ?
          $""""
               <SplitHorizontal>1545</SplitHorizontal>
               <TopRowBottomPane>1</TopRowBottomPane>
               <SplitVertical>26820</SplitVertical>
               <LeftColumnRightPane>13</LeftColumnRightPane>
               <ActivePane>2</ActivePane>
               <Panes>
                <Pane>
                 <Number>3</Number>
                </Pane>
                <Pane>
                 <Number>1</Number>
                </Pane>
                <Pane>
                 <Number>2</Number>
                 <ActiveCol>1</ActiveCol>
                 <ActiveRow>{initialRow}</ActiveRow>
                </Pane>
                <Pane>
                 <Number>0</Number>
                 <ActiveRow>59</ActiveRow>
                 <ActiveCol>16</ActiveCol>
                </Pane>
               </Panes>
               """" :
           $""""
               <SplitHorizontal>1545</SplitHorizontal>
               <TopRowBottomPane>{initialRow}</TopRowBottomPane>
               <ActivePane>2</ActivePane>
               <Panes>
                <Pane>
                 <Number>3</Number>
                 <ActiveRow>1</ActiveRow>
                 <ActiveCol>1</ActiveCol>
                </Pane>
                <Pane>
                 <Number>2</Number>
                 <ActiveRow>{initialRow}</ActiveRow>
                 <ActiveCol>1</ActiveCol>
                </Pane>
               </Panes>
               <SplitHorizontal>1545</SplitHorizontal>
               <TopRowBottomPane>{initialRow}</TopRowBottomPane>
               <ActivePane>2</ActivePane>
               <Panes>
                <Pane>
                 <Number>3</Number>
                 <ActiveRow>1</ActiveRow>
                 <ActiveCol>1</ActiveCol>
                </Pane>
                <Pane>
                 <Number>2</Number>
                 <ActiveRow>{initialRow}</ActiveRow>
                 <ActiveCol>1</ActiveCol>
                </Pane>
               </Panes>
               """") +
           $""""
               <ProtectObjects>False</ProtectObjects>
               <ProtectScenarios>False</ProtectScenarios>
              </WorksheetOptions>
              <AutoFilter x:Range="R1C1:R{rowCount}C{columnCount}"
               xmlns="urn:schemas-microsoft-com:office:excel">
              </AutoFilter>
             </Worksheet>
             """";
    }

    static List<string> MakeSymbolListForMasterRegex(SortedDictionary<string, SymbolInList> masterSymbolListForWorksheet)
    {
        List<string> symbolListForMasterRegex = new();

        foreach (var symbol in masterSymbolListForWorksheet.Keys.Where(s => !string.IsNullOrEmpty(s)))
            symbolListForMasterRegex.Add(symbol);
        return symbolListForMasterRegex;
    }

    static readonly Regex patIbdLiveWatchReadyList = new Regex(@"^(IBD Live)\s*(Ready|Watch)\.csv");
    static readonly Regex patLargeMidSmallCap = new Regex(@"^(Large|Mid|Small) Cap\.csv");
    static readonly Regex patStartsWithIbd = new Regex(@"^IBD");
    static readonly Regex patIPO = new Regex(@"IPO");
    static readonly Regex patSchwabAccountPositions = new Regex(@"^(sss) (.*)\.csv$");

    SortedDictionary<string, SymbolInList> MakeMasterSymbolListForWorksheet(SortedDictionary<string, DateTime> mapFileNameToMostRecentFileDate, bool debug)
    {
        SortedDictionary<string, SymbolInList> masterSymbolListForWorksheet = new();
        foreach ((var fileName, var mapSymbolToDatesToMetrics) in mapFileNameToSymbolToDatesToAttributes)
        {
            foreach ((var symbol, var mapDatesToMetrics) in mapSymbolToDatesToMetrics)
            {
                if (debug) WriteLine($"file={fileName} Adding {symbol} to comparisonSheet for {string.Join(",", mapDatesToMetrics.Keys.OrderByDescending(x => x).Select(x => x.Date.ToString("yy-MMM-dd-ddd")).Take(5))}");
                var mostRecentDateForThisFile = mapDatesToMetrics.Keys.Max(x => x.Date);
                if (mostRecentDateForThisFile == mapFileNameToMostRecentFileDate[fileName].Date)
                {
                    if (masterSymbolListForWorksheet.ContainsKey(symbol))
                    {
                        masterSymbolListForWorksheet[symbol].Lists.Add(fileName);
                    }
                    else
                    {
                        var tmp = new SortedSet<string>();
                        tmp.Add(fileName);
                        masterSymbolListForWorksheet.Add(symbol, new SymbolInList { Name = symbol, Lists = tmp });
                    }
                }
            }
        }
        return masterSymbolListForWorksheet;
    }

    /// <summary>
    /// Conditionally Adds a sort order prefix to a file name based on specific patterns. These will be stripped off when converting the file name to a column header name.
    /// </summary>
    /// <param name="fileName">The input string representing the file name to potentially modify.</param>
    /// <returns>The modified file name with an added sort order prefix if applicable.</returns>
    private static string MaybeAddSortOrderPrefix(string fileName)
    {
        if (patIbdLiveWatchReadyList.Match(fileName).Success)
            // add digits to sort order prefix for Schwab account position to order columns with in the group of columns for the Schwab account positions
            fileName = patIbdLiveWatchReadyList.Replace(fileName, m => $"uuu {m.Groups[1].Value} {m.Groups[2].Value}");
        else if (patLargeMidSmallCap.Match(fileName).Success)
            fileName = patLargeMidSmallCap.Replace(fileName, m => $"qqq {m.Groups[1].Value} Cap");
        else if (fileName.Contains("RS Line"))
            fileName = "bbb " + fileName; // add a sort order prefix for RS Line files
        else if (patStartsWithIbd.Match(fileName).Success)
            fileName = "ttt " + fileName; // add a sort order prefix for IBD files
        else if (patIPO.Match(fileName).Success)
            fileName = "iii " + fileName; // add a sort order prefix for IPO files
        else if (fileName == "buy.csv")
            fileName = "www buy.csv"; // group this next to the watch list 
        else if (patSchwabAccountPositions.Match(fileName).Success)
        {
            fileName = patSchwabAccountPositions.Replace(fileName, m => $@"{m.Groups[1].Value}000 {m.Groups[2].Value}");
        }
        return fileName;
    }

    static string[] MakeFileNamesArray(SortedDictionary<string, int> mapFileNameToColumnPosition)
    {
        var columnCount = 0;
        var fileNames = new string[mapFileNameToColumnPosition.Count];
        foreach (var fn in mapFileNameToColumnPosition.Keys) fileNames[columnCount++] = fn; // deep copy
        return fileNames;
    }
    static bool isScreenTip(string attributeName, Dictionary<string, (bool tip, string displayName, Func<string, string> convert)> mapAttributeNameToScreenTip)
    {
        var screenTip = false;
        if (!mapAttributeNameToScreenTip.ContainsKey(attributeName))
        {
            screenTip = true;
        }
        else
        {
            screenTip = mapAttributeNameToScreenTip[attributeName].tip;
        }
        //WriteLine($"attribute={attributeName} tip={screenTip}");
        return screenTip;
    }
    void GenerateXMLWorksheetDataRows(int BACK_HISTORY_COUNT, 
        //AutoMultiDimSortedDictionary<string, AutoMultiDimSortedDictionary<string, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string, string>>>> mapFileNameToSymbolToDatesToAttributes, 
        SortedDictionary<string, SymbolInList> masterSymbolListForWorksheet, 
        //SortedDictionary<string, int> mapSortOrderPrefixedFileNameToColumnPosition, 
        //SortedDictionary<string, string> mapFileNameToFileNameNoExt, 
        AutoMultiDimSortedDictionary<string, AutoInitSortedDictionary<string, SortedSet<DateTime>>> mapSymbolToFileNameToDates, 
        StringBuilder sbXMLWorksheetRows,
        StringBuilder sbCSV, bool debug, ref int rowCount, ref int columnCurrent, 
        AutoInitSortedDictionary<string, ExcelStyle> excelStyles, 
        SortedSet<string> fileNamesWithHistory, int columnCount, 
        DateTime[] historyDirectoryDateArray, 
        Dictionary<string, (bool tip, string displayName, Func<string, string> convert)> mapAttributeNameToScreenTip, 
        (string orderBy, string colorCodedAttributeName) workSheet)
    {
        foreach (var equitySymbol in masterSymbolListForWorksheet.Keys.Where(s => !string.IsNullOrEmpty(s))) // for each row (one equity symbol per row)
        {
            var symbol= equitySymbol;
            columnCurrent = 0;

            sbXMLWorksheetRows.AppendLine("""  <Row ss:AutoFitHeight="0">"""); // begin row
            sbXMLWorksheetRows.AppendLine($"""    <Cell><Data ss:Type="Number">{rowCount + 1}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>""");
            bool skipOverBlankColumn = false;
            AutoMultiDimSortedDictionary<string/*fileName*/, AutoMultiDimSortedDictionary<string/*symbol*/, AutoInitSortedDictionary<string, string>>> latestAttributes = new();
            var latestFileNameForThisSymbol = string.Empty;
            //var latestFileNamesSetsForThisSymbol = new DataFileCatagories(new SortedSet<string>(), new SortedSet<string>(), new SortedSet<string>());

            // Each input data file contains two columns: the first column is the
            // symbol name and the second column is an attribute of
            // that symbol such as the current price or average true
            // value. The header row for the column tells us the
            // attribute name. The attribute name must be contained in
            // the file name. This file has been read into a
            // dictionary that becomes manifest as a single
            // spreadsheet column.


            //foreach (var fileName in DescendingComparer<string>.reverse ? masterSymbolListForWorksheet[symbol].Lists.Reverse() : masterSymbolListForWorksheet[symbol].Lists) // for each column
            //SortedSet<string> filesWithSymbolsInThisRow = new();
            //masterSymbolListForWorksheet[symbol].Lists.ToList().ForEach(fileName => filesWithSymbolsInThisRow.Add(mapFileNameToSortOrderPrefixedFileName[fileName]));
            //foreach ((var sortOrderPrefixedFileName,var fileName) in mapSortOrderPrefixedFileNameToFileName)
            //foreach(var sortOrderPrefixedFileName in filesWithSymbolsInThisRow)
            foreach(var fileName in masterSymbolListForWorksheet[symbol].Lists.OrderBy(fn => mapFileNameToSortOrderPrefixedFileName[fn]))
            {
                //var fileName = mapSortOrderPrefixedFileNameToFileName[sortOrderPrefixedFileName];
                latestFileNameForThisSymbol = fileName;
                var fileNameWithSortOrderPrefix = mapFileNameToSortOrderPrefixedFileName[fileName];// MaybeAddSortOrderPrefix(fileName);
                var index = mapSortOrderPrefixedFileNameToColumnPosition[fileNameWithSortOrderPrefix]; // @@todo@@ abandon mapSortOrderPrefixedFileNameToColumnPosition and just use the order of the keys in mapSortOrderPrefixedFileNameToFileName and increment index every iteration
                if (debug) WriteLine($"fileName={fileName} index={index}");
                while (columnCurrent < index)
                {
                    if (/*columnCurrent > 0 &&*/ columnCurrent < columnCount)
                    {
                        skipOverBlankColumn = true;
                        //sbXMLRows.AppendLine($"""    <Cell><Data ss:Type="String"></Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""");
                    }
                    columnCurrent++;
                }
                if (columnCurrent >= 0 && columnCurrent < columnCount)
                {
                    var skipToIndex = skipOverBlankColumn ? $" ss:Index=\"{columnCurrent + 2}\"" : "";
                    skipOverBlankColumn = false;
                    var matchNotes = patNotes.Match(fileName);
                    var matchUnrealizedGains = patUnrealizedGains.Match(fileName);
                    var matchGroups = patGroup.Match(fileName);
                    var matchGroupSymbol = patSymbolName.Match(fileName);
                    var matchDecoratedSymbol = patDecoratedSymbol.Match(fileName);
                    var matchAtr21Day = patAtr21Day.Match(fileName);
                    var matchNetLiqPercent = patNetLiqPercent.Match(fileName);
                    var matchAccounts = patAccount.Match(fileName);
                    var equityAttributes = mapFileNameToSymbolToDatesToAttributes[fileName][symbol];
                    var attributeKeys = equityAttributes.Keys;
                    var latest = attributeKeys.OrderByDescending(d => d).ToArray<DateTime>()[0];
                    if (matchNotes.Success)
                    {
                        var name = matchNotes.Groups[2].Value;
                        var formulas = matchNotes.Success ? matchNotes.Groups[3].Value == "Formulas" : false;
                        var nakedSymbol = symbol;
                        if (mapFileNameToSymbolToDatesToAttributes[fileName].ContainsKey(symbol))
                            nakedSymbol = symbol;
                        else
                            nakedSymbol = symbol.Contains(":") ? symbol.Split(':')[^1] : symbol; // accommodate symbols from ExtractTOSStockGroups.py that prefix the symbol with the account name and equity group name seperated by colons
                        var stock = mapFileNameToSymbolToDatesToAttributes[fileName][nakedSymbol][latest];
                        if(stock.ContainsKey($"{name}Notes"))
                        {
                            var notes = stock[$"{name}Notes"]; // notes may contain a single comment, a single formula, or multiple notes
                            //var daysOld = stock.ContainsKey($"{name}DaysOld") ? stock[$"{name}DaysOld"] : "99";
                            notes = patDateTime_ddd_MMM_dd_YYYY_ddd_MMM_dd_YYYY.Replace(notes, m => m.Groups[6].Value); // remove the date of data entry and leave the date of the data (report)   
                            notes = notes.Trim();
                            var line = formulas ?
                                $"""    <Cell{skipToIndex} ss:Formula="{notes}" ss:StyleID="s70"><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""" 
                                : 
                                $"""    <Cell{skipToIndex}><ss:Data ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40">{notes}</ss:Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""" ;
                            sbXMLWorksheetRows.AppendLine(line);
                        }
                    }
                    else if (matchGroupSymbol.Success)
                    {
                       var stock = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                       var symbolName = stock["SymbolName"];
                       var line = $"""    <Cell{skipToIndex}><Data ss:Type="String">{symbolName}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""";
                       sbXMLWorksheetRows.AppendLine(line);
                    }
                    else if (matchAccounts.Success)
                    {
                        var stock = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                        var account = stock["Account"];
                        var line = $"""    <Cell{skipToIndex}><Data ss:Type="String">{account}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""";
                        sbXMLWorksheetRows.AppendLine(line);
                    }
                    else if (matchAtr21Day.Success)
                    {
                        var stock = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                        var name = matchAtr21Day.Groups[3].Value;
                        var atr21day = stock[name];
                        var line = $"""    <Cell{skipToIndex} ss:StyleID="s75Percent"><Data ss:Type="Number">{atr21day}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""";
                        sbXMLWorksheetRows.AppendLine(line);
                    }
                    else if (matchNetLiqPercent.Success)
                    {
                        var stock = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                        var name = matchNetLiqPercent.Groups[3].Value;
                        var netliqpercent = stock[name];
                        var line = $"""    <Cell{skipToIndex} ss:StyleID="s75Percent"><Data ss:Type="Number">{netliqpercent}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""";
                        sbXMLWorksheetRows.AppendLine(line);
                    }
                    else if (matchGroups.Success)
                    {
                        var stock = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                        var name = matchGroups.Groups[3].Value;
                        var groupOrManagedBy = stock[name];
                        var line = $"""    <Cell{skipToIndex}><Data ss:Type="String">{groupOrManagedBy}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""";
                        sbXMLWorksheetRows.AppendLine(line);
                    }
                    else if (matchUnrealizedGains.Success)
                    {
                        var name = matchUnrealizedGains.Groups[3].Value; // use the 3rd capture group in the file name to know the column name to fetch
                        var stock = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                        var sharesOrProfitLossOrUnitCost = stock[name].Trim();
                        var style = "";
                        var dailyChange = patDailyChange.Match(name).Success;
                        if ("ProfitLoss" == name || "PortionOfTotalAccount" == name || dailyChange || name == "IRR")
                        {
                            if (sharesOrProfitLossOrUnitCost.IndexOf("$") >= 0)
                            {
                                sharesOrProfitLossOrUnitCost = sharesOrProfitLossOrUnitCost.Replace("$", "");
                            }
                            var val = +(double.Parse((sharesOrProfitLossOrUnitCost=="-"|| sharesOrProfitLossOrUnitCost == "N/A") ?"0": sharesOrProfitLossOrUnitCost) / 100);
                            sharesOrProfitLossOrUnitCost = val.ToString();
                            if (val < -0.08)
                            {
                                style = " ss:StyleID=\"s73\"";
                            }
                            else if (val < 0)
                            {
                                style = " ss:StyleID=\"s71\"";
                            }
                            else
                            {
                                style = " ss:StyleID=\"s70\"";
                            }
                        }
                        else if ("NetLiquidValue" == name || "Price" == name || "ProfitLossDollar" == name)
                        {
                            double val = ParseDouble(sharesOrProfitLossOrUnitCost, fileName);
                            if (val < 0)
                                style = " ss:StyleID=\"s74\"";
                            else
                                style = " ss:StyleID=\"s72\"";
                        }
                        else {
                            style = " ss:StyleID=\"s76Comma\""; 
                        }
                        var line = $"""    <Cell{skipToIndex}{style}><Data ss:Type="Number">{sharesOrProfitLossOrUnitCost}</ss:Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""";
                        sbXMLWorksheetRows.AppendLine(line);
                    }
                    else
                    {  
                        var displaySymbol = symbol;
                        if (matchDecoratedSymbol.Success)
                        {
                            var stock = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                            displaySymbol = stock["DecoratedSymbol"];  // decorated symbols are color coded and have a prefix ageMetric so you can sort by the ageMetric (such as days until earnings)
                        }
                        // The most recent csv files are in the Download directory.
                        // Older copies of the files are in sub-directories of the Download directory whose names contains the date.
                        // For the first time thru the loop we fetch the most recent file and that is not in a history directory yet so we don't compare the dates (yet).
                        // If we forgot to download a file in the past, we assume the current symbol was supposed to be in that file and don't advance kkk so we can compare it again.
                        var age = 0;
                        var jjj = 0;
                        var kkk = 0;
                        var count = 0;
                        var datesForThisSymbol = mapFileNameToSymbolToDatesToAttributes[fileName][symbol].Keys.OrderByDescending(d => d).ToArray<DateTime>();
                        var datesForThisFileNameOnly = mapSymbolToFileNameToDates[symbol][fileName];
                        var datesForThisSymbolArray = datesForThisSymbol.ToArray();
                        var DatesForThisSymbolCount = datesForThisSymbol.Length;
                        while (kkk < DatesForThisSymbolCount && jjj < historyDirectoryDateArray.Length)
                        {
                            var historyDirectoryDate = historyDirectoryDateArray[jjj++].Date;
                            var symbolDate = datesForThisSymbolArray[kkk++].Date;
                            var eq = historyDirectoryDate == symbolDate;
                            if (count == 0 && !eq)
                            {
                                age++;
                                jjj--;
                            }
                            else if (eq)
                            {
                                age++;
                            }
                            else
                            {
                                /*
                                if (!datesForThisFileNameOnly.Contains(historyDirectoryDate)) // we missed a downloading a file for this day. Assume the symbol was there.
                                {
                                    // bug: symbol date is 6/30/2023 and we historyDirectoryDate is 12/1/2023
                                    if (orderBy == "age" && symbol == "SPXC" && fileName == "RS Line Blue Dot.csv")
                                        WriteLine($"We missed downloading a file symbol={symbol} age={age} hd={historyDirectoryDate} dtfs={symbolDate}");
                                    age++;
                                    kkk--; // back up and look for a match on the next go around.
                                }
                                else
                                {
                                    break;
                                }*/
                                break;
                            }
                            count++;
                        }
                        AutoInitSortedDictionary<string, string> attributeTable = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                        latestAttributes[symbol][fileName] = attributeTable;
                        var screenTip = string.Join(",", attributeTable
                            .Where(e => !string.IsNullOrEmpty(e.Value) && e.Key != "seq" && isScreenTip(e.Key, mapAttributeNameToScreenTip))
                            .Select(e => mapAttributeNameToScreenTip.ContainsKey(e.Key) ? $"""{mapAttributeNameToScreenTip[e.Key].displayName}={patRemoveLeadingZeros.Replace(mapAttributeNameToScreenTip[e.Key].convert(e.Value), m => m.Groups[2].Value)}""" : $"""{e.Key}={e.Value}"""));
                        screenTip = $"age={age}," + screenTip;
                        screenTip = TruncateString(screenTip);
                        var styleName2 = "s68";
                        if (age != 0 /*&& fileNamesWithHistory.Contains(fileName)*/)
                        {
                            var ageMetric = 1.0 * age / (BACK_HISTORY_COUNT + 1);
                            ageMetric = Math.Log(ageMetric * (BACK_HISTORY_COUNT + 1)) / Math.Log(BACK_HISTORY_COUNT + 1);
                            // @@todo@@: replace this with a Dictionary<string, Func<AutoInitSortedDictionary<string, string>, string>> so we can add new color coding metrics easily
                            if (workSheet.colorCodedAttributeName == "Ind Group Rank") ColorCodeBy_Ind_Group_Rank(attributeTable);
                            else if (workSheet.colorCodedAttributeName == "Comp Rating") ColorCodeBy_Comp_Rating(attributeTable, "Comp Rating");
                            else if (workSheet.colorCodedAttributeName == "RS Rating") ColorCodeBy_Comp_Rating(attributeTable, "RS Rating");
                            else if (workSheet.colorCodedAttributeName == "EPS Rating") ColorCodeBy_Comp_Rating(attributeTable, "EPS Rating");
                            else if (workSheet.colorCodedAttributeName == "Up/Down Vol") ColorCodeHueByMetric(attributeTable, "Up/Down Vol");  //ColorCodeBy_UpDown_Rating(attributeTable, "Up/Down Vol");
                            else if (workSheet.colorCodedAttributeName == "ROE") ColorCodeHueByMetric(attributeTable, "ROE");
                            else if (workSheet.colorCodedAttributeName == "Price % Chg") ColorCodeHueByMetric(attributeTable, "Price % Chg");
                            else if (workSheet.colorCodedAttributeName == "Daily Closing Range") ColorCodeHueByMetric(attributeTable, "Daily Closing Range");
                            else if (workSheet.colorCodedAttributeName == "Weekly Closing Range") ColorCodeHueByMetric(attributeTable, "Weekly Closing Range");
                            else if (workSheet.colorCodedAttributeName == "Market Cap (mil)") ColorCodeBy_MarketCap(attributeTable, "Market Cap (mil)");
                            else if (string.IsNullOrEmpty(workSheet.colorCodedAttributeName)) ColorCodeBy_Dollar_Volume(attributeTable);
                            else ColorCodeHueByMetric(attributeTable, workSheet.colorCodedAttributeName);
                            stockExcelSaturationAgeStyle.InputMetric = ageMetric;
                            var RGBHexColor = stockExcelSaturationAgeStyle.ColorHexRGB; // add pattern here for dollar volume
                            ExcelStyle style;
                            if (attributeTable.ContainsKey("50-Day Avg $ Vol (1000s)") && !string.IsNullOrEmpty(attributeTable["50-Day Avg $ Vol (1000s)"]))
                            {
                                // Flag stocks with low dollar volume by using a pattern
                                double dollarVolume = 0;
                                if (double.TryParse(attributeTable["50-Day Avg $ Vol (1000s)"].Replace(",", ""), out dollarVolume))
                                {
                                    dollarVolume = dollarVolume * 1000;
                                }
                                else
                                {
                                    dollarVolume = 0;
                                    //WriteLine($"Error: {fileName} {symbol} 50-Day Avg $ Vol (1000s)={attributeTable["50-Day Avg $ Vol (1000s)"]}");
                                }
                                if (dollarVolume < 10 * 1000 * 1000)
                                {
                                    var pattern = "ThinVertStripe";
                                    style = new ExcelStyle { Color = RGBHexColor, Name = "s" + RGBHexColor + pattern, Pattern = pattern, PatternColor = "000000" };
                                    excelStyles[RGBHexColor + pattern] = style;
                                    styleName2 = "s" + RGBHexColor + pattern;
                                    //style = new ExcelStyle { Color = RGBHexColor, Name = "s" + RGBHexColor };
                                    //excelStyles[RGBHexColor] = style;
                                    //styleName2 = "s" + RGBHexColor;
                                    screenTip = PrependToScreenTip(symbol, screenTip);
                                }
                                else if (dollarVolume < 20 * 1000 * 1000)
                                {
                                    var pattern = "ThinHorzStripe";
                                    style = new ExcelStyle { Color = RGBHexColor, Name = "s" + RGBHexColor + pattern, Pattern = pattern, PatternColor = "000000" };
                                    excelStyles[RGBHexColor + pattern] = style;
                                    styleName2 = "s" + RGBHexColor + pattern;
                                    //style = new ExcelStyle { Color = RGBHexColor, Name = "s" + RGBHexColor };
                                    //excelStyles[RGBHexColor] = style;
                                    //styleName2 = "s" + RGBHexColor;
                                    screenTip = PrependToScreenTip(symbol, screenTip);
                                }
                                else
                                {
                                    style = new ExcelStyle { Color = RGBHexColor, Name = "s" + RGBHexColor };
                                    excelStyles[RGBHexColor] = style;
                                    styleName2 = "s" + RGBHexColor;
                                }
                            }
                            else
                            {
                                style = new ExcelStyle { Color = RGBHexColor, Name = "s" + RGBHexColor };
                                excelStyles[RGBHexColor] = style;
                                styleName2 = "s" + RGBHexColor;
                            }
                            if (debug1) WriteLine($"fileName={fileName} symbol={symbol} age={age} ageMetric={ageMetric}");
                        }
                        var seqNoAndSymbol = "";
                        if (fileName.Contains("IBD 50 Index")) // special case because this list is always in the same order
                            seqNoAndSymbol = $"{mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest]["seq"]} {displaySymbol}";
                        else if (string.IsNullOrEmpty(workSheet.orderBy) || workSheet.orderBy.ToLower() == "symbol")
                            seqNoAndSymbol = displaySymbol;
                        else if (workSheet.orderBy == "age")
                            seqNoAndSymbol = $"{age.ToString("000")} {displaySymbol}";
                        else
                        {
                            //System.Diagnostics.Debugger.Break(); // this is a bug, we should never get here
                            var attrValue = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest][workSheet.orderBy];
                            if (!string.IsNullOrEmpty(attrValue) && mapAttributeNametoAttributeAttributes.ContainsKey(workSheet.orderBy) && mapAttributeNametoAttributeAttributes[workSheet.orderBy].numeric && mapAttributeNametoAttributeAttributes[workSheet.orderBy].orderByConvert != null)
                            {
                                attrValue = mapAttributeNametoAttributeAttributes[workSheet.orderBy].orderByConvert(attrValue, workSheet.orderBy);
                                seqNoAndSymbol = $"{attrValue} {displaySymbol}";
                            }
                            else if (Int32.TryParse(attrValue ?? "999", out int attr1))
                                seqNoAndSymbol = $"{attr1.ToString("000")} {displaySymbol}";
                            else if (double.TryParse(attrValue ?? "  0.00", out double attr2))
                                seqNoAndSymbol = $"{attr2.ToString((attr2 < 0 ? "" : " ") + "   0.00")} {displaySymbol}";
                            else
                                seqNoAndSymbol = $"999 {displaySymbol}";
                        }
                        //if (orderBy == "age" && symbol == "SPXC" && fileName == "RS Line Blue Dot.csv")  WriteLine($"symbol={symbol} age={age} seqNoAndSymbol={seqNoAndSymbol} attributes={attributes}");
                        var HRef = fileName.Contains("Group.") ? "" : $"ss:HRef=\"https://marketsmith.investors.com/mstool?Symbol={displaySymbol}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0\"";
                        var cell = $"""    <Cell{skipToIndex} ss:StyleID="{styleName2}" {HRef} x:HRefScreenTip="{screenTip}"><Data ss:Type="String">{seqNoAndSymbol}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}--> """;
                        sbXMLWorksheetRows.AppendLine(cell);
                    }
                }
                columnCurrent++;
            }

            // add other stock metrics here. Unlike the other spreadsheet columns these do not come from a single file.            

            rowCount++;
            columnCurrent = /*mapFileNameToFileNameNoExt*/this.mapFileNameToColumnHeader.Keys.Count;
            var metricCells = ""; skipOverBlankColumn = true;
            //var symbolOnly = symbol.Split(':').Reverse().FirstOrDefault(); // remove the Think or Swim equity/stock group  prefix

            AutoInitSortedDictionary<string, string> attributes = FindAttributesForSymbol(symbol, latestAttributes);

            // Fix for CS7036: Ensure the correct arguments are passed to the 'colorCodeHue' Action delegate.

            foreach ((var position, (var attributeName, var numeric, var screenTip, var displayName, var convert, _, _, _, var colorCodeHue)) in mapPositionToAttributeName)
            {
                var styleName = "s68";
                var skipToIndex = skipOverBlankColumn ? $" ss:Index=\"{columnCurrent + 2}\"" : "";
                var type = numeric ? "Number" : "String";
                var value = convert(attributes[attributeName]);
                if (!string.IsNullOrEmpty(value))
                {
                    value = patRemoveFinalM.Replace(value, m => m.Groups[1].Value);
                    value = patRemoveLeadingZeros.Replace(value, m => m.Groups[2].Value);
                }
                // Ensure the 'colorCodeHue' Action is invoked with the correct arguments
                if (colorCodeHue != null && numeric && value!= "-")
                {
                    colorCodeHue(this, attributes, attributeName); 
                    stockExcelSaturationAgeStyle.InputMetric = 0.6; // reset the input metric for the next color coding. This is the saturation.
                    var RGBHexColor = stockExcelSaturationAgeStyle.ColorHexRGB;
                    if (!generateMasterRegexList && debug1)
                    {
                        WriteLine($"fileName={latestFileNameForThisSymbol} symbol={symbol} attributeName={attributeName} value={value} RGBHexColor={RGBHexColor} stockExcelSaturationAgeStyle=({stockExcelSaturationAgeStyle})");
                    }

                    var style = new ExcelStyle { Color = RGBHexColor, Name = "s" + RGBHexColor };
                    excelStyles[RGBHexColor] = style;
                    styleName = "s" + RGBHexColor;
                }

                metricCells += $"""    <Cell{skipToIndex} ss:StyleID="{styleName}"><Data ss:Type="{type}">{value}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>""";
                columnCurrent++;
                skipOverBlankColumn = false;
            }
            sbXMLWorksheetRows.AppendLine(metricCells);
            sbXMLWorksheetRows.AppendLine($"  </Row> <!-- {rowCount} -->");
            sbXMLWorksheetRows.AppendLine();
        }

        static double ParseDouble(string sharesOrProfitLossOrUnitCost, string fileName)
        {
            try
            {
                return double.Parse(sharesOrProfitLossOrUnitCost);
            }
            catch (Exception ex)
            {
                WriteLine($"file = {fileName} exception: {ex}");
                throw ex;
            }
        }
    }

    private static AutoInitSortedDictionary<string, string> FindAttributesForSymbol(string symbol, AutoMultiDimSortedDictionary<string, AutoMultiDimSortedDictionary<string, AutoInitSortedDictionary<string, string>>> latestAttributes)
    {
        var maxAttributeCount = 0;
        var attributes = new AutoInitSortedDictionary<string, string>();
        //string bestFile = String.Empty;
        foreach (var fileName in latestAttributes[symbol].Keys)
        {
            var attributeCount = latestAttributes[symbol][fileName].Count;
            if (attributeCount > maxAttributeCount)
            {
                maxAttributeCount = latestAttributes[symbol][fileName].Count;
                attributes = latestAttributes[symbol][fileName]; // use the attributes from the file with the most attributes
                                                                 //bestFile = fileName;
            }
            if (debug1) WriteLine($"fileName={fileName} attributeCount = {attributeCount} attributes={string.Join(",", attributes.Keys)}");
        }

        return attributes;
    }

    private static string PrependToScreenTip(string? symbol, string screenTip)
    {
        screenTip = symbol + " " + screenTip;
        screenTip = TruncateString(screenTip);
        return screenTip;
    }

    private static string TruncateString(string screenTip, int maxLength=255)
    { 
        var maxTries = 20;
        for(var count=0; count < maxTries && screenTip.Length > maxLength; count++)
        {
            var lastComma= screenTip.LastIndexOf(',');
            screenTip = screenTip.Substring(0, lastComma > -1 ? lastComma: maxLength);
        }
        if (screenTip.Length > maxLength)
            screenTip = screenTip.Substring(0, maxLength);
        return screenTip;
    }

    void ColorCodeBy_Dollar_Volume(AutoInitSortedDictionary<string, string> attributeTable)
    {
        var dollarVolume = attributeTable["50-Day Avg $ Vol (1000s)"];
        if (string.IsNullOrEmpty(dollarVolume) || dollarVolume == "-")
        {
            //WriteLine($"File name = {fileName} {symbol} missing dollar vol");
            stockExcelSaturationAgeStyle.Hue = 290.0 / 360.0 * 255.0; // 290 is purple for "missing value" or "no value" or "N/A"
        }
        else
        {
            var dollarVol = double.Parse(dollarVolume.Replace(",", "")) * 1000;
            if (dollarVol > 20e6)
                stockExcelSaturationAgeStyle.Hue = 120.0 / 360.0 * 255.0; // green
            else if (dollarVol > 15e6)
                stockExcelSaturationAgeStyle.Hue = 90.0 / 360.0 * 255.0;
            else if (dollarVol > 10e6)
                stockExcelSaturationAgeStyle.Hue = 60.0 / 360.0 * 255.0; // yellow
            else if (dollarVol > 5e6)
                stockExcelSaturationAgeStyle.Hue = 30.0 / 360.0 * 255.0;
            else
                stockExcelSaturationAgeStyle.Hue = 0.0 / 360.0 * 255.0; // red
        }
    }
    void ColorCodeBy_Ind_Group_Rank(AutoInitSortedDictionary<string, string> attributeTable)
    {
        var indGroupRankString = attributeTable["Ind Group Rank"];
        if (string.IsNullOrEmpty(indGroupRankString) || indGroupRankString == "-")
        {
            //WriteLine($"File name = {fileName} {symbol} missing dollar vol");
            stockExcelSaturationAgeStyle.Hue = 290.0 / 360.0 * 255.0;
        }
        else
        {
            var indGroupRank = int.Parse(indGroupRankString.Trim());
            if (indGroupRank < 40)
                stockExcelSaturationAgeStyle.Hue = 120.0 / 360.0 * 255.0; // green
            else if (indGroupRank < 80)
                stockExcelSaturationAgeStyle.Hue = 90.0 / 360.0 * 255.0;
            else if (indGroupRank < 120)
                stockExcelSaturationAgeStyle.Hue = 60.0 / 360.0 * 255.0; // yellow
            else if (indGroupRank < 160)
                stockExcelSaturationAgeStyle.Hue = 30.0 / 360.0 * 255.0;
            else
                stockExcelSaturationAgeStyle.Hue = 0.0 / 360.0 * 255.0; // red
        }
    }

    void ColorCodeBy_Comp_Rating(AutoInitSortedDictionary<string, string> attributeTable, string metric)
    {
        var compRatingString = attributeTable[metric];
        if (string.IsNullOrEmpty(compRatingString) || compRatingString == "-")
        {
            //WriteLine($"File name = {fileName} {symbol} missing dollar vol");
            stockExcelSaturationAgeStyle.Hue = 290.0 / 360.0 * 255.0;
        }
        else
        {
            // 99->120/360*255==85
            // 0->0
            //
            // 85/99==0.85859
            var compRating = int.Parse(compRatingString.Trim());
            stockExcelSaturationAgeStyle.Hue = compRating * 0.85859;
        }
    }
    void ColorCodeBy_UpDown_Rating(AutoInitSortedDictionary<string, string> attributeTable, string metricName)
    {
        var upDownVolStr = attributeTable[metricName];
        if (string.IsNullOrEmpty(upDownVolStr) || upDownVolStr == "-")
        {
            stockExcelSaturationAgeStyle.Hue = 290.0 / 360.0 * 255.0;
        }
        else
        {
            var upDownVol = double.Parse(upDownVolStr.Trim());
            var min = doubleMin[metricName];
            var max = doubleMax[metricName];
            if (mapAttributeNametoAttributeAttributes[metricName].max != double.MinValue)
            {
                max = Math.Min(max, mapAttributeNametoAttributeAttributes[metricName].max);
            }
            upDownVol = (upDownVol - min) / (max - min);
            stockExcelSaturationAgeStyle.Hue = upDownVol * 120.0 / 360.0 * 255.0;
        }
    }
    /// <summary>
    /// Adjusts the hue and the saturation of a color based on a specified ageMetric from a dictionary, indicating data status visually.
    /// For example, the Saturation metric could be the age of the data in days (the number of days that equity has been on a watch list), and the hue could be closing range or the price change.
    /// Hues are in the range of 0 to 180, where 0 is red (undesirable), 62 is yellow (warning), 120 is green (good), and 180 is blue (best).
    /// Output: stockExcelSaturationAgeStyle.Hue
    /// </summary>
    /// <param name="attributeTable">Provides a mapping of ageMetric names to their corresponding string values for hue calculation.</param>
    /// <param name="metricName">Specifies the key used to retrieve the ageMetric value from the provided mapping.</param>
    void ColorCodeHueByMetric(AutoInitSortedDictionary<string, string> attributeTable, string metricName)
    {
        var metricStr = attributeTable[metricName];
        if (string.IsNullOrEmpty(metricStr) || metricStr == "-")
        {
            stockExcelSaturationAgeStyle.Hue = 290.0 / 360.0 * 255.0; // this is purple indicating no data or not applicable
        }
        else
        {
            double metric;
            try
            {
                metric = double.Parse(metricStr.Trim());
            }
            catch(Exception ex)
            {
                WriteLine($"Error parsing metric '{metricStr}' for metricName '{metricName}': {ex.Message}");
                throw ex; // rethrow the exception to handle it upstream
            }
            double min;
            try
            {
                min = doubleMin[metricName];
            }
            catch (Exception ex)
            {
                WriteLine($"Error accessing doubleMin for metricName '{metricName}': {ex.Message}");
                throw ex; // rethrow the exception to handle it upstream
            }
            double max;
            try
            {
                max = doubleMax[metricName];
            }
            catch (Exception ex)
            {
                WriteLine($"Error accessing doubleMax for metricName '{metricName}': {ex.Message}");
                throw ex; // rethrow the exception to handle it upstream
            }
            if (mapAttributeNametoAttributeAttributes[metricName].max != double.MinValue)
            {
                max = Math.Min(max, mapAttributeNametoAttributeAttributes[metricName].max);
            }
            if (mapAttributeNametoAttributeAttributes[metricName].min != double.MaxValue)
            {
                min = Math.Max(min, mapAttributeNametoAttributeAttributes[metricName].min);
            }
            metric = Math.Min(metric, max);
            metric = Math.Max(metric, min);
            metric = (metric - min) / (max - min);
            var hue = metric * 180.0 / 360.0 * 255.0;  // 180 is blue, the maximum hue value for blue in the HSL color model. Usually blue is good. 0 is red which is the worst. 
            stockExcelSaturationAgeStyle.Hue = hue;
            // WriteLine($" upDownVol={upDownVolStr} => upDownVol={upDownVol} => stockExcelSaturationAgeStyle.Hue={stockExcelSaturationAgeStyle.Hue}");
        }
    }
    void ColorCodeHueByEarningsSalesMetric(AutoInitSortedDictionary<string, string> attributeTable, string metricName)
    {
        var metricStr = attributeTable[metricName];
        if (string.IsNullOrEmpty(metricStr) || metricStr == "-")
        {
            stockExcelSaturationAgeStyle.Hue = 290.0 / 360.0 * 255.0; // this is purple indicating no data or not applicable
        }
        else
        {
            double metric;
            try
            {
                metric = double.Parse(metricStr.Trim());
            }
            catch (Exception ex)
            {
                WriteLine($"Error parsing metric '{metricStr}' for metricName '{metricName}': {ex.Message}");
                throw ex; // rethrow the exception to handle it upstream
            }
            double min;
            try
            {
                min = doubleMin[metricName];
            }
            catch (Exception ex)
            {
                WriteLine($"Error accessing doubleMin for metricName '{metricName}': {ex.Message}");
                throw ex; // rethrow the exception to handle it upstream
            }
            if (mapAttributeNametoAttributeAttributes[metricName].min != double.MaxValue)
            {
                min = Math.Max(min, mapAttributeNametoAttributeAttributes[metricName].min);
            }
            double max;
            try
            {
                max = doubleMax[metricName];
            }
            catch (Exception ex)
            {
                WriteLine($"Error accessing doubleMax for metricName '{metricName}': {ex.Message}");
                throw ex; // rethrow the exception to handle it upstream
            }
            if (mapAttributeNametoAttributeAttributes[metricName].max != double.MinValue)
            {
                max = Math.Min(max, mapAttributeNametoAttributeAttributes[metricName].max);
            }
            if(metric > 0)
            {
                var logMetric = Math.Log10(metric);
                if (logMetric > 2 && logMetric < 3)
                {
                    logMetric = 0; // don't allow negative values
                }
                var logMax = Math.Log10(max);
            }
            else
            {
                metric = Math.Min(metric, max);
            }
            metric = Math.Min(metric, max);
            metric = Math.Max(metric, min);
            metric = (metric - min) / (max - min);
            var hue = metric * 180.0 / 360.0 * 255.0;  // 180 is blue, the maximum hue value for blue in the HSL color model. Usually blue is good. 0 is red which is the worst. 
            stockExcelSaturationAgeStyle.Hue = hue;
            // WriteLine($" upDownVol={upDownVolStr} => upDownVol={upDownVol} => stockExcelSaturationAgeStyle.Hue={stockExcelSaturationAgeStyle.Hue}");
        }
    }
    void ColorCodeBy_MarketCap(AutoInitSortedDictionary<string, string> attributeTable, string metricName)
    {
        var marketCapStr = attributeTable[metricName];
        if (string.IsNullOrEmpty(marketCapStr) || marketCapStr == "-")
        {
            stockExcelSaturationAgeStyle.Hue = 290.0 / 360.0 * 255.0;
        }
        else
        {
            var marketCap = Log10(double.Parse(marketCapStr.Trim()));
            var min = doubleMin[metricName];
            min = min>0?Log10(min):0;
            var max = doubleMax[metricName];
            max = Log10(max);
            if (mapAttributeNametoAttributeAttributes[metricName].max != double.MinValue)
            {
                max = Math.Min(max, mapAttributeNametoAttributeAttributes[metricName].max);
            }
            if (mapAttributeNametoAttributeAttributes[metricName].min != double.MaxValue)
            {
                min = Math.Max(min, mapAttributeNametoAttributeAttributes[metricName].min);
            }
            marketCap = Math.Min(marketCap, max);
            marketCap = Math.Max(marketCap, min);
            marketCap = (marketCap - min) / (max - min);
            stockExcelSaturationAgeStyle.Hue = marketCap * 180.0 / 360.0 * 255.0;
            // WriteLine($" upDownVol={upDownVolStr} => upDownVol={upDownVol} => stockExcelSaturationAgeStyle.Hue={stockExcelSaturationAgeStyle.Hue}");
        }
    }
    static readonly Regex patUnitCost = new Regex(@"(UnitCost)", RegexOptions.Compiled);
    static readonly Regex patProfitLossDollar = new Regex(@"(ProfitLossDollar)", RegexOptions.Compiled);
    static readonly Regex patProfitLoss = new Regex(@"(ProfitLoss)", RegexOptions.Compiled);
    static readonly (Regex regex, string replacement)[] readableColumnNames = new (Regex, string)[]
    {
        (patUnitCost, "Unit Cost"),
        (patProfitLossDollar, "Profit Loss $"),
        (patProfitLoss, "Profit Loss %"),
        (new Regex(@"DATR21ManagedByGroupPercent"), "21 Day $ ATR Managed By Group %"),
        (new Regex(@"DATR21ManagedByPercent", RegexOptions.Compiled), "21 Day $ ATR Managed By %"),
        (new Regex(@"DATR21Percent", RegexOptions.Compiled), "21 Day $ ATR %"),
        (new Regex(@"DATR21AccntGroupPercent", RegexOptions.Compiled), "21 Day $ ATR Account Group %"),
        (new Regex(@"DATR21AccntPercent", RegexOptions.Compiled), "21 Day $ ATR Account %"),
        (new Regex(@"DecoratedSymbol", RegexOptions.Compiled), "Decorated Symbol"),
        (new Regex(@"Schwab SymbolName", RegexOptions.Compiled), "Schwab Symbol Name"),
        (new Regex(@"Schwab NetLiquidValue", RegexOptions.Compiled), "Schwab Net Liquid $"),
        (new Regex(@"Schwab NetLiqPercent", RegexOptions.Compiled), "Schwab Net Liquid %"),
        (new Regex(@"AccntType", RegexOptions.Compiled), "Account Type"),
        (new Regex(@"Schwab ATR21Day", RegexOptions.Compiled), "Schwab 21 Day ATR"),
        (new Regex(@"Schwab DATR21", RegexOptions.Compiled), "Schwab 21 Day $ ATR"),
        (new Regex(@"Schwab TotalCost", RegexOptions.Compiled), "Schwab Total Cost $"),
        (new Regex(@"DATR21Percent", RegexOptions.Compiled), "21 Day $ ATR %"),
    };

    static SortedDictionary<string, string> MakeMapFileNameToColumnHeader(string[] fileNames)
    {
        var mapFileNameToFileNameNoExt = new SortedDictionary<string, string>();
        foreach (var fileName in fileNames)
        {   //@@todo@@ Fix this and check for successful pattern match. Also: insert some spaces between words in the column name so excel can wrap them.
            var fileNameWithoutOptionalPrefix = patFileNameOrderPrefix.Match(fileName).Groups[3].Value; // strip off the prefix that implements the sorted order
            var match = patFileExtension.Match(fileNameWithoutOptionalPrefix);
            var fileNameWithOutExtension = match.Success ? match.Groups[1].Value : Path.GetFileName(fileNameWithoutOptionalPrefix);
            foreach ((var regex, var replacement) in readableColumnNames)
            {
                // Replace the file names with more readable column names
                if (regex.Match(fileNameWithOutExtension).Success)
                {
                    //if (!generateMasterRegexList) WriteLine($"Replacing {fileNameWithOutExtension} with {replacement}");
                    fileNameWithOutExtension = regex.Replace(fileNameWithOutExtension, replacement);
                }
                else if (!generateMasterRegexList && fileNameWithOutExtension.Contains("Schwab"))
                {
                    //WriteLine($"Not replacing {fileNameWithOutExtension} with {replacement}");
                }
            }

            mapFileNameToFileNameNoExt.Add(fileName, fileNameWithOutExtension);
        }
        return mapFileNameToFileNameNoExt;
    }
    int GenerateXMLWorksheetColumnDeclarations(//SortedDictionary<string, int> mapSortOrderPrefixedFileNameToColumnPosition, SortedDictionary<string, string> mapFileNameToColumnHeader 
        StringBuilder sbXMLWorksheetRows, 
        StringBuilder sbCSV, 
        List<string> xmlColumnWidths/*, 
        string[] sortOrderPrefixedFileNames*/)
    {
        var columnCount = 0;
        foreach (var sortOrderPrefixedFileName in sortOrderPrefixedFileNames)
        {
            var fileNameWithSortOrderPrefix = sortOrderPrefixedFileName;
            mapSortOrderPrefixedFileNameToColumnPosition[fileNameWithSortOrderPrefix] = columnCount++; // populate mapSortOrderPrefixedFileNameToColumnPosition here! @@todo@@ this is superfluous, remove it.
            var fileNameWithoutOptionalPrefix = patFileNameOrderPrefix.Match(sortOrderPrefixedFileName).Groups[3].Value;
            var match = patFileExtension.Match(fileNameWithoutOptionalPrefix);
            var fileNameWithOutExtension = match.Success ? match.Groups[1].Value : Path.GetFileName(sortOrderPrefixedFileName);
            var m = patNotes.Match(sortOrderPrefixedFileName);
            var comments = m.Success ? m.Groups[3].Value == "Comments" : false;
            var formulas = m.Success ? m.Groups[3].Value == "Formulas" : false;
            if (m.Success && !(comments|formulas))
            {
                xmlColumnWidths.Add($"""<Column ss:Width="{NoteColumnWidth}" />""");
            }
            else
            {
                xmlColumnWidths.Add($"""<Column ss:Width="{SymbolColumnWidth}" />""");
            }
        }
        sbXMLWorksheetRows.AppendLine();
        return columnCount;
    }

    static void UpdateMapSymbolToFileNameToDates(AutoMultiDimSortedDictionary<string, AutoInitSortedDictionary<string, SortedSet<DateTime>>> mapSymbolToFileNameToDates, string fileName, AutoMultiDimSortedDictionary<string, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string, string>>> csvData)
    {
        foreach (var sym in csvData.Keys)
        {
            foreach (var date in csvData[sym].Keys)
            {
                // Why is this necessary?
                if (!mapSymbolToFileNameToDates[sym].ContainsKey(fileName))
                    mapSymbolToFileNameToDates[sym][fileName] = new SortedSet<DateTime>();
                mapSymbolToFileNameToDates[sym][fileName].Add(date);
            }
        }
    }

    static void UpdateMapFileNameToSymbolToDateToAttributes(AutoMultiDimSortedDictionary<string/*file name*/, AutoMultiDimSortedDictionary<string /*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*ageMetric name*/, string/*ageMetric value*/>>>> mapFileNameToSymbolToDatesToAttributes,
                                                 string fileName,
                                                 AutoMultiDimSortedDictionary<string, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string, string>>> csvData)
    {
        foreach (var symbolKey in csvData.Keys)
        {
            foreach (var dateKey in csvData[symbolKey].Keys)
            {
                // Why is this necessary?
                //if (!mapFileNameToSymbolToDatesToAttributes[fileName].ContainsKey(fileName)) mapFileNameToSymbolToDatesToAttributes[fileName][symbolKey] = new AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string, string>>();
                //mapFileNameToSymbolToDatesToAttributes[fileName][symbolKey] = csvData[symbolKey][dateKey];
                foreach (var attributeNameKey in csvData[symbolKey][dateKey].Keys)
                    mapFileNameToSymbolToDatesToAttributes[fileName][symbolKey][dateKey][attributeNameKey] = csvData[symbolKey][dateKey][attributeNameKey];

            }
        }
    }
    static Regex patDateTime_YYYY_MMM_dd_ddd = new Regex(@"^(([0-9]+)-(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-([0-9]+))-(Mon|Tue|Wed|Thu|Fri|Sat|Sun)$");
    static Regex patDate_YYYY_MM_dd = new Regex(@"^(\d{4})-(\d{2})-(\d{2})");
    static Regex patDateTime_ddd_MMM_dd_YYYY_ddd_MMM_dd_YYYY = new Regex(@"^(Mon|Tue|Wed|Thu|Fri|Sat|Sun) ((Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) ([0-9]+),? ([0-9]+)):\s*&lt;((Mon|Tue|Wed|Thu|Fri|Sat|Sun) ((Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) ([0-9]+),? ([0-9]+)))&gt;");
    static Regex patDateTime_ddd_MMM_dd_YYYY = new Regex(@"^(Mon|Tue|Wed|Thu|Fri|Sat|Sun) ((Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) ([0-9]+),? ([0-9]+))");

    /// <summary>
    /// Extracts a DateTime value from a specified file path based on its directory structure and naming convention.
    /// If there is no Date embedded in the path, it returns the last write date-time of the file.
    /// </summary>
    /// <param name="filePath">The path to the file from which the date and time information will be extracted.</param>
    /// <returns>Returns a DateTime object representing either the parsed date from the directory name or the last write time of
    /// the file.</returns>
    static DateTime ExtractDateTimeFromFilePath(string filePath)
    {
        var dirs = filePath.Split('\\');
        if (dirs.Length > 3)
        {
            var dir = dirs[4];
            if (dir.Length > 12 && patDateTime_YYYY_MMM_dd_ddd.IsMatch(dir))
            {
                var s = dir.Substring(0, dir.Length - 4); // remove the day of week from end
                var date = DateTime.Parse(s);
                return date;
            }
            else
            {
                return new FileInfo(Environment.ExpandEnvironmentVariables(filePath)).LastWriteTime;
            }
        }
        else
        {
            return new FileInfo(Environment.ExpandEnvironmentVariables(filePath)).LastWriteTime;
        }
    }
    static void MakeHistoryDateSet(IEnumerable<History> history, DateTime defaultDateTime, ref SortedSet<DateTime> historyDates)
    {
        foreach (var h in history)
        {
            var date = ExtractDateTimeFromFilePath(h.Name);
            historyDates.Add(date);
        }
    }

    static string MakeFileNamePretty(string fileName)
    {
        var fileNameWithoutOptionalPrefix = patFileNameOrderPrefix.Match(fileName).Groups[3].Value;
        var match = patFileExtension.Match(fileNameWithoutOptionalPrefix);
        var fileNameWithOutExtension = match.Success ? match.Groups[1].Value : Path.GetFileName(fileName);
        return fileNameWithOutExtension;
    }
    static string FormatInteger(string strValue, string _)
    {
        if (Int32.TryParse(strValue, out Int32 value))
            if (value < 0)
                return "999";
            else
                return value.ToString("000");
        else
            return strValue;
    }
    static string MapLogDoubleToInteger(string strValue, string attrName)
    {
        var min = doubleMin[attrName];
        if(min>0)
            min = Math.Log10(min);
        var max = doubleMax[attrName];
        if(max>0)
            max = Math.Log10(max);
        if (double.TryParse(strValue, out double value))
        {
            if(value >0)
                value = Math.Log10(value);
            if (value == max)
            {
                return "999";
            }
            else
            {
                var mappedValue = (value - min) / (max - min);
                var result = (int)(mappedValue * 1000 + 0.5);
                return result.ToString("000");
            }
        }
        return "999";
    }

    // Fix for CS0305: Correcting the generic type usage for AutoInitIntSortedDictionary

    // The issue arises because the method `ColorCodeByHue` is using `AutoInitIntSortedDictionary<string, string>`,
    // but the type `AutoInitIntSortedDictionary<K>` only accepts one type argument. 
    // Based on the context, it seems the intended type should be `AutoInitIntSortedDictionary<string>`.

    static void ColorCodeByHue3(CompareTickerSymbolListsInCSVFilesMainProgram main, CompareTickerSymbolListsInCSVFiles.AutoInitSortedDictionary<string,string> attributes, string metric)
    {
        main.ColorCodeHueByMetric(attributes, metric);
    }

    // Add new "orderbys" here
    static Dictionary<int, AttributeAttributes> mapPositionToAttributeName = new Dictionary<int, AttributeAttributes>{
        //                            name                      numeric screenTip displayName        convert orderByConvert
        { 0, new AttributeAttributes("ROE"                             , true , true , "ROE"                             , e=>e                      , FormatFloat, 100,-100) },
        { 1, new AttributeAttributes("ROE 5-Yr Avg"                    , true , false, "ROE 5-Yr Avg"                    , e=>e                      , FormatFloat, 100,-100) },
        { 2, new AttributeAttributes("21 Day ATR %"                    , true , false, "21 Day ATR %"                    , e=>e                      , (e,_)=>{ if(string.IsNullOrEmpty(e)) e= ""; else if(e!="-") e= ""+(Double.Parse(e)*10);  return FormatPercentFloat(e,_); }, 100, 0) }, // instead of displaying a decimal point which takes up room, just multiply by 10 instead
        { 3, new AttributeAttributes("Earnings Stability"              , true , false, "Earnings Stability"              , e=>e                      , FormatInteger, 99,0) },
        { 4, new AttributeAttributes("Sustainable Growth %"            , true , false, "Sustainable Growth %"            , e=>e                      , FormatFloat, 650, 0, ColorCodeByHue3) },
        { 5, new AttributeAttributes("EPS % Chg Lst Yr"                , true , false, "EPS % Chg Lst Yr"                , e=>e                      , FormatFloat, 100, 0, ColorCodeByHue3) },
        { 6, new AttributeAttributes("EPS % Chg 1 Yr Ago"              , true , false, "EPS % Chg 1 Yr Ago"              , e=>e                      , FormatFloat, 5000,0, ColorCodeByHue3) },
        { 7, new AttributeAttributes("EPS % Growth 1 Yr"               , true , false, "EPS % Growth 1 Yr"               , e=>e                      , FormatFloat, 5000,0, ColorCodeByHue3) },
        { 8, new AttributeAttributes("EPS % Growth 3 Yr"               , true , false, "EPS % Growth 3 Yr"               , e=>e                      , FormatFloat, 5000,0, ColorCodeByHue3) },
        { 9, new AttributeAttributes("EPS % Growth 5 Yr"               , true , false, "EPS % Growth 5 Yr"               , e=>e                      , FormatFloat, 5000,0, ColorCodeByHue3) },
        {10, new AttributeAttributes("3-Yr EPS Growth Geq 5-Yr"        , false, false, "3-Yr EPS Growth Geq 5-Yr"        , e=>e                      , null) },
        {11, new AttributeAttributes("EPS % Growth 5 Yr Pct Rnk"       , true , false, "EPS % Growth 5 Yr Pct Rnk"       , e=>e                      , FormatFloat, 99,0, ColorCodeByHue3) },
        {12, new AttributeAttributes("EPS Accel 3 Qtrs"                , false, false, "EPS Accel 3 Qtrs"                , e=>e                      , null) },
        {13, new AttributeAttributes("EPS % Chg Last Qtr"              , true , false, "EPS % Chg Last Qtr"              , e=>e                      , FormatFloat, 1000, 0, ColorCodeByHue3) },
        {14, new AttributeAttributes("EPS % Chg 1 Q Ago"               , true , false, "EPS % Chg 1 Q Ago"               , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {15, new AttributeAttributes("EPS % Chg 2 Q Ago"               , true , false, "EPS % Chg 2 Q Ago"               , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {16, new AttributeAttributes("EPS % Chg 3 Q Ago"               , true , false, "EPS % Chg 3 Q Ago"               , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {17, new AttributeAttributes("Avg EPS % Chg 2Q"                , true , false, "Avg EPS % Chg 2Q"                , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {18, new AttributeAttributes("Avg EPS % Chg 3Q"                , true , false, "Avg EPS % Chg 3Q"                , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {19, new AttributeAttributes("Avg EPS % Chg 4Q"                , true , false, "Avg EPS % Chg 4Q"                , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {20, new AttributeAttributes("Avg EPS % Chg 5Q"                , true , false, "Avg EPS % Chg 5Q"                , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {21, new AttributeAttributes("Avg EPS % Chg 6Q"                , true , false, "Avg EPS % Chg 6Q"                , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {22, new AttributeAttributes("EPS % Chg Lst Q Gtr 3-Yr Growth" , false, false, "EPS % Chg Lst Q Gtr 3-Yr Growth" , e=>e                      , null) },
        {23, new AttributeAttributes("EPS Trailing 4 Qtrs"             , true , false, "EPS Trailing 4 Qtrs"             , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {24, new AttributeAttributes("EPS Trl 4Q Geq EPS Lst Fiscal Yr", false, false, "EPS Trl 4Q Geq EPS Lst Fiscal Yr", e=>e                      , null) },
        {25, new AttributeAttributes("EPS Lst Yr Gtr EPS 4 Yrs Ago"    , false, false, "EPS Lst Yr Gtr EPS 4 Yrs Ago"    , e=>e                      , null) },
        {26, new AttributeAttributes("EPS Est Cur Qtr %"               , true , false, "EPS Est Cur Qtr %"               , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {27, new AttributeAttributes("EPS Est Cur Yr %"                , true , false, "EPS Est Cur Yr %"                , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {28, new AttributeAttributes("EPS Est Next Yr %"               , true , false, "EPS Est Next Yr %"               , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {29, new AttributeAttributes("EPS % Chg Last Qtr (-/+)"        , true , false, "EPS % Chg Last Qtr (-/+)"        , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {30, new AttributeAttributes("EPS % Chg 1 Q Ago (-/+)"         , true , false, "EPS % Chg 1 Q Ago (-/+)"         , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {31, new AttributeAttributes("EPS % Chg 2 Q Ago (-/+)"         , true , false, "EPS % Chg 2 Q Ago (-/+)"         , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {32, new AttributeAttributes("EPS % Chg 3 Q Ago (-/+)"         , true , false, "EPS % Chg 3 Q Ago (-/+)"         , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {33, new AttributeAttributes("EPS Surprise"                    , true , false, "EPS Surprise"                    , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {34, new AttributeAttributes("Fiscal EPS Lst Yr"               , true , false, "Fiscal EPS Lst Yr"               , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {35, new AttributeAttributes("Fiscal EPS 1 Yr Ago"             , true , false, "Fiscal EPS 1 Yr Ago"             , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {36, new AttributeAttributes("Fiscal EPS 2 Yrs Ago"            , true , false, "Fiscal EPS 2 Yrs Ago"            , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {37, new AttributeAttributes("Fiscal EPS 3 Yrs Ago"            , true , false, "Fiscal EPS 3 Yrs Ago"            , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {38, new AttributeAttributes("Fiscal EPS 4 Yrs Ago"            , true , false, "Fiscal EPS 4 Yrs Ago"            , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {39, new AttributeAttributes("Fiscal EPS 6 Yrs Ago"            , true , false, "Fiscal EPS 6 Yrs Ago"            , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {40, new AttributeAttributes("Fiscal EPS 5 Yrs Ago"            , true , false, "Fiscal EPS 5 Yrs Ago"            , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {41, new AttributeAttributes("Symbol"                          , false, false, "Symbol"                          , e=>e                      , null) },
        {42, new AttributeAttributes("Sales % Chg Lst Qtr"             , true , false, "Sales % Chg Lst Qtr"             , e=>e                      , FormatFloat, 999,0, (main,attributes,metric)=>{ main.ColorCodeHueByMetric(attributes,metric); }) },
        {43, new AttributeAttributes("Avg Sales % Chg 2Q"              , true , false, "Avg Sales % Chg 2Q"              , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {44, new AttributeAttributes("Avg Sales % Chg 3Q"              , true , false, "Avg Sales % Chg 3Q"              , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {45, new AttributeAttributes("Avg Sales % Chg 4Q"              , true , false, "Avg Sales % Chg 4Q"              , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {46, new AttributeAttributes("Avg Sales % Chg 5Q"              , true , false, "Avg Sales % Chg 5Q"              , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {47, new AttributeAttributes("Avg Sales % Chg 6Q"              , true , false, "Avg Sales % Chg 6Q"              , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {48, new AttributeAttributes("Sales Accel 2 Qtrs"              , false, false, "Sales Accel 2 Qtrs"              , e=>e                      , null) },
        {49, new AttributeAttributes("Sales Accel 3 Qtrs"              , false, false, "Sales Accel 3 Qtrs"              , e=>e                      , null) },
        {50, new AttributeAttributes("Sales % Chg Lst Yr"              , true , false, "Sales % Chg Lst Yr"              , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {51, new AttributeAttributes("Sales Growth 3 Yr"               , true , false, "Sales Growth 3 Yr"               , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {52, new AttributeAttributes("Sales Growth 5 Yr"               , true , false, "Sales Growth 5 Yr"               , e=>e                      , FormatFloat, 999,0, ColorCodeByHue3) },
        {53, new AttributeAttributes("Symbol-2"                        , false, false, "Symbol"                          , e=>e                      , null) },
        {54, new AttributeAttributes("Annual Sales (mil)"              , true , false, "Annual Sales (mil)"              , e=>e                      , FormatFloat, 999,0) },
        {55, new AttributeAttributes("seq"                             , true , false, "seq"                             , e=>e                      , null) },
        {56, new AttributeAttributes("Current Price"                   , true , true , "Price"                           , e=>e                      , null) },
        {57, new AttributeAttributes("Price % Chg"                     , true , true , "Price % Chg"                     , e=>e                      , MapLogDoubleToInteger,5,-5)},
        {58, new AttributeAttributes("Comp Rating"                     , true , true , "Comp Rating"                     , e=>e                      , FormatInteger) },
        {59, new AttributeAttributes("EPS Rating"                      , true , true , "EPS Rating"                      , e=>e                      , null ) },
        {60, new AttributeAttributes("RS Rating"                       , true , true , "RS Rating"                       , e=>e                      , FormatInteger) },
        {61, new AttributeAttributes("A/D Rating"                      , false, true , "A/D Rating"                      , e=>e                      , null) },
        {62, new AttributeAttributes("SMR Rating"                      , false, true , "SMR Rating"                      , e=>e                      , null) },
        {63, new AttributeAttributes("50-Day Avg $ Vol (1000s)"        , true , true , "50-Day Avg $M Vol"               , e=>FormatDollarVolume(e,"M"), (e,_) => FormatDollarVolume(e))},
        {64, new AttributeAttributes("Ind Group Rank"                  , true , true , "Ind Group Rank"                  , e=>e                      , FormatInteger) },
        {65, new AttributeAttributes("Industry Name"                   , false, true , "Industry Name"                   , e=>e                      , null) },
        {66, new AttributeAttributes("Sector"                          , false, false, "Sector"                          , e=>e                      , null) },
        {67, new AttributeAttributes("Name"                            , false, true , "Name"                            , e=>e                      , null) },
        {68, new AttributeAttributes("Sponsor Rating"                  , false, false, "Sponsor Rating"                  , e=>e                      , null) },
        {69, new AttributeAttributes("Funds % Increase"                , true , false, "Funds % Increase"                , e=>e                      , FormatFloat, 10,0) },
        {70, new AttributeAttributes("Number of Funds"                 , true , false, "Number of Funds"                 , e=>e                      , FormatInteger, 2000, 0) },
        {71, new AttributeAttributes("Funds %"                         , true , false, "Funds %"                         , e=>e                      , FormatFloat, 99, 0) },
        {72, new AttributeAttributes("Mgmt %"                          , true , false, "Mgmt %"                          , e=>e                      , FormatFloat, 99, 0) },
        {73, new AttributeAttributes("% Off High"                      , true , false, "% Off High"                      , e=>e                      , FormatFloat, 3, -20) },
        {74, new AttributeAttributes("EPS Due Date"                    , false, false, "EPS Due Date"                    , e=>e                      , null) },
        {75, new AttributeAttributes("Days to Earnings"                , true , false, "Days to Earnings"                , e=>e                      , null) },
        {76, new AttributeAttributes("Up/Down Vol"                     , true , true , "Up/Down Vol"                     , e=>e                      , FormatFloat, 2) },
        {77, new AttributeAttributes("Yield"                           , true , true , "Yield"                           , e=>e                      , FormatFloat) },
        {78, new AttributeAttributes("ETF"                             , false, false, "ETF"                             , e=>e                      , (e,n)=>e) },
        {79, new AttributeAttributes("Daily Closing Range"             , true , false, "Daily Closing Range"             , e=>e                      , FormatPercentFloat, 100, 0) },
        {80, new AttributeAttributes("Weekly Closing Range"            , true , false, "Weekly Closing Range"            , e=>e                      , FormatPercentFloat, 100, 0) },
        {81, new AttributeAttributes("Market Cap (mil)"                , true , true , "Market Cap (mil)"                , e=>e                      , MapLogDoubleToInteger, 4_000_000.4,2.0) },
        {82, new AttributeAttributes("30 Day ATR %"                    , true , false, "30 Day ATR %"                    , e=>e                      , FormatPercentFloat, 100, 0) },
        {83, new AttributeAttributes("50 Day ATR %"                    , true , false, "50 Day ATR %"                    , e=>e                      , FormatPercentFloat, 100, 0) },
        {84, new AttributeAttributes("Symbol-3"                        , false, false, "Symbol"                          , e=>e                      , null) },
        {85, new AttributeAttributes("Account"                         , false, false, "Account"                         , e=>e                      , null) },
        {86, new AttributeAttributes("Group"                           , false, false, "Group"                           , e=>e                      , null) },
        {87, new AttributeAttributes("Cost"                            , true , false, "Cost"                            , e=>e                      , FormatFloat) },
        {88, new AttributeAttributes("Total Cost"                      , true , false, "Total Cost"                      , e=>e                      , FormatFloat) }

    };
    static string FormatDollarVolume(string e, string M ="")
    {
        if ( string.IsNullOrEmpty(e) || e == "-")
            return e;
        else 
            return $"{double.Parse(e.Replace(",", "")) / 1e3:000.00}{M}";
    }
    static string FormatFloat(string e, string _)
    {
        if (string.IsNullOrEmpty(e) || e == "-") 
        { 
            //return e;
            return "  999.99";
        } 
        else {
            var val = double.Parse(e.Replace(",", ""));
            return val < 0 ? $"{val:000.00}" : $"{val:  000.00}"; 
        } 
    }
    static string FormatPercentFloat(string e, string _)
    {
        if (string.IsNullOrEmpty(e) || e == "-")
        {
            return e;
        }
        else
        {
            var val = double.Parse(e.Replace(",", ""));
            if (val == 100) val = 99.99; // conserve space and don't use three digits
            var result = $"{val:00}";
            if (result == "100")
                result = "99";
            else if (result == "999")
                result = "99";
            //var result = val < 0 ? $"{val:00}" : $"{}";
            return result;
        }
    }

    static Dictionary<string, AttributeAttributes> MakeMapAttributeNametoAttributeAttributes(Dictionary<int, AttributeAttributes> mapPositionToAttributeName)
    {
        var result = new Dictionary<string, AttributeAttributes>();
        foreach (var (_, attributeAttributes) in mapPositionToAttributeName)
        {
            result.Add(attributeAttributes.name, attributeAttributes);
        }
        return result;
    }
    static Dictionary<string, AttributeAttributes> mapAttributeNametoAttributeAttributes = MakeMapAttributeNametoAttributeAttributes(mapPositionToAttributeName);
    static Regex patRestoreCommas = new Regex(@"\&comma;");
    static Regex patRestoreNewLines = new Regex(@"\&(newline|nl);");
    static Regex patTag = new Regex(@"<B>\s*<Font[^>]*>\s*![A-Z][A-Z0-9\.=]*!\s*</Font>\s*</B>");
    static (Regex regex, string replacement)[] patHighLightKeyWord = {
        (new Regex("NEUTRAL"), """<B><Font html:Size="14" html:Color="#FFC000">NEUTRAL</Font></B>"""),
        //(new Regex("(?<!STRONG\\s*)SELL"),    """<B><Font html:Size="14" html:Color="#FF0000">SELL</Font></B>"""),
        (new Regex("Reiterate Buy",RegexOptions.IgnoreCase), """<B><Font html:Color="#008080">Reiterate Buy</Font></B>"""),
        //(new Regex("(?<!STRONG\\s*)BUY"),"""<B><Font html:Size="14" html:Color="#008080">BUY</Font></B>"""),
        (new Regex("LOW"),"""<B><Font html:Size="14" html:Color="#008080">LOW</Font></B>"""),
        (new Regex("HIGH"),"""<B><Font html:Size="14" html:Color="#FF0000">HIGH</Font></B>"""),
        (new Regex("(!AVOID!|!SHORT!|!SELLSIG!)"),"""<B><Font html:Size="14" html:Color="#FF0000">$1</Font></B>"""), 
        (new Regex("(!HOLD!|!SELL!|!WED!|!BSW!|!SL!|!WFV!|!GSW!|!HBZ!|!HTF!|!DDTL!|!W4BASE!|!TOPPING!)"),"""<B><Font html:Size="14" html:Color="#1873BD">$1</Font></B>"""), // !WED! == Watch Every Day! SL! == Short List BSW! == Best Stocks to Watch! WFV=Mike Webster's Favorite
        (new Regex("(!BUY!|!ACTIONABLE!|!BREAKINGOUT!)"),"""<B><Font html:Size="14" html:Color="#19BD19">$1</Font></B>"""), // !WED! == Watch Every Day! SL! == Short List BSW! == Best Stocks to Watch! WFV=Mike Webster's Favorite
        (new Regex("(!3DE!|!3DS!)"),"""<B><Font html:Size="14" html:Color="#FF5500">$1</Font></B>"""), 
        (new Regex("(!BP([=0-9\\.]+)?!)"),"""<B><Font html:Size="14" html:Color="#3EB534">$1</Font></B>"""), // Green !BP! == at buy point      
        (new Regex("(!W4PB!|!BOW!)"),"""<B><Font html:Size="14" html:Color="#00CCFF">$1</Font></B>"""), // W4PB=Wait for pull back BOW=Break out watch
        (new Regex("STRONG\\s*BUY"),"""<B><I><Font html:Size="15" html:Color="#008080">STRONG BUY</Font></I></B>"""),
        (new Regex("STRONG\\s*SELL"),"""<B><I><Font html:Size="15" html:Color="#FF0000">STRONG SELL</Font></I></B>"""),
        (new Regex("MODERATE"),"""<B><Font html:Size="14" html:Color="#00CCFF">MODERATE</Font></B>"""),
        (new Regex("MEDIUM"),"""<B><Font html:Size="14" html:Color="#00CCFF">MEDIUM</Font></B>"""),
        //(new Regex("HOLD"),"""<B><Font html:Size="14" html:Color="#00CCFF">HOLD</Font></B>"""),     
        (new Regex(@"\b(PO)\b:?"),"""<B><Font html:Size="14" html:Color="#00CCFF">PT:</Font></B>"""),
        (new Regex(@"(FMV|Fair Value Estimate|PRICE TARGET|TARGET PRICE)(:)?",RegexOptions.IgnoreCase),"""<B><Font html:Size="14" html:Color="#00CCFF">PT$2</Font></B>"""),
        (new Regex(@"(Target Price:?|Price Target( Raised to)?:?|PRICE TARGET:|(Raising )?Fair Value Estimate( to)?:?|) (\$[0-9\.]+)"),"""<B><Font html:Size="14" html:Color="#00CCFF">$0</Font></B>""")
    };
    static SortedDictionary<string,string> flags = new SortedDictionary<string, string> {
        {"!WED!", "Watch Every Day"},
        {"!BSW!", "Best Stocks to Watch"},
        {"!SL!", "Short List"},
        {"!3DE!", "Triple digit earnings or earnings estimates" },
        {"!3DS!", "Triple digit sales" },
        {"!WFV!", "Mike Webster's Favorite"},
        {"!BP!", "At Buy Point"},
        {"!BOW!", "Break Out Watch"},
        {"!AVOID!", "Avoid buying or owning"},
        {"!DDTL!", "Draw down trend line" },
        {"!BREAKINGOUT!", "breaking out now" },
        {"!SHORT!", "Short"},
        {"!BUY!", "buy" },
        {"!HBZ!", "Holding in buy zone" },
        {"!HTF!", "High Tight Flag" },
        {"!SELL!", "sell" },
        {"!SELLSIG!", "sell signal" },
        {"!ACTIONABLE!", "actionable" },
        {"!MODERATE!", "moderate" },
        {"!TOPPING!", "Topping?" },
        {"!MEDIUM!", "medium" },
        {"!STRONGBUY!", "strong buy" },
        {"!STRONGSELL!", "strong sell" },
        {"!HOLD!", "hold" },
        {"!W4PB!", "Wait for Pull Back"},
        {"!WF4B", "Wait for base to build" },
        {"!GSW!","IBD50 Growth Stocks to Watch" }
    };

    static int allowOldDataFilesHoursOld = ComputeAllowedOldDataFilesHours();//-13;  // use old csv data files up to 13 hours after midnight before coding resulting excel spreadsheet cell as purple to indicate old or missing data

    static int ComputeAllowedOldDataFilesHours()
    {
        int result = 13; // default to 13 hours old because market closes at 13 hours after midnight. Allow use to use the default data file until 1pm the next day.
        // if today is Saturday subtract 24 hours
        if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
        {
            result = 24;
        }else if(DateTime.Now.DayOfWeek == DayOfWeek.Sunday)
            {
            // if today is Sunday subtract 48 hours
            result = 48;
        }
        return result;
    }
    static Regex patNoteHeaderBody = new Regex(@"^(\d{2}\s+(Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d+,?\s+\d{2}\s*<B>[A-Z]+</B>:)(\s*(\s*![A-Z]+!\s*))*(.*)$");
    static Regex patTags = new Regex(@"(\s*(![A-Z]+!))+");
    AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*ageMetric name*/, string/*ageMetric value*/>>> ParseCSV(
        string fileName,
        DateTime currentDataFileDateTime,
        AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*ageMetric name*/, string/*ageMetric value*/>>>? defaultValues,
        int maxEventAge)
    {
        (var baseNameNoVersionFn, var versioned) = GetBaseFileNameNoVersions(Path.GetFileName(fileName));
        var annotation = versioned ? $"({baseNameNoVersionFn}) " : "";
        var useCfg = true;
        AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*ageMetric name*/, string/*ageMetric value*/>>> rows = new();
        var configuration = new CsvHelper.Configuration.CsvConfiguration(new CultureInfo("en-US", false))
        {
            MissingFieldFound = null,
            HasHeaderRecord = true
        };       
        //WriteLine($"ParseCSV: open {fileName}");
        //System.Diagnostics.Debugger.Break();
        // if(!generateMasterRegexList) WriteLine($"ParseCSV: open {fileName}");
        using (StreamReader reader = new StreamReader(fileName)) // use the raw file to open the csv file
        //using (var csv = new CsvHelper.CsvReader(reader,  new CsvHelper.Configuration.CsvConfiguration {  Delimiter = ",",  HasHeaderRecord = true })) {
        //using (CsvReader csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        using (CsvReader csv = useCfg ? new CsvReader(reader, configuration) : new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            //csv.Configuration.HasHeaderRecord = false;

            var isNotesFileName = patNotes.Match(baseNameNoVersionFn);
            var isCommentsFileName = isNotesFileName.Success ? isNotesFileName.Groups[3].Value == "Comments" : false;
            var isFormulasFileName = isNotesFileName.Success ? isNotesFileName.Groups[3].Value == "Formulas" : false;
            var isSymbolFileName = patSymbolName.Match(baseNameNoVersionFn);
            var isUnrealizedGainName = patUnrealizedGains.Match(baseNameNoVersionFn);
            var isSchwabGroupFileName = patGroup.Match(baseNameNoVersionFn);
            var isSchwabATR21DayFileName = patAtr21Day.Match(baseNameNoVersionFn);
            var isSchwabNetLiqPercentFileName = patNetLiqPercent.Match(baseNameNoVersionFn);
            var isAccountFileName = patAccount.Match(baseNameNoVersionFn);
            var isDecoratedSymbolFileName = patDecoratedSymbol.Match(baseNameNoVersionFn);
            //isAccountFileName = new Regex(@$"({patSortOrderPrefix})\s*([a-zA-Z0-9]*)\s*Account(\.csv)?$").Match(baseNameNoVersionFn);
            var result = new AutoInitSortedDictionary<string, string>();
            var count = 0;
            while (csv.Read())
            {
                var skip = false;
                if (count++ > 0)
                {
                    var name = "";
                    var notes = "";
                    var symbol = csv.GetField<string>("Symbol");
                    symbol = symbol is null ? "" : symbol.Trim();
                    var symbolOnly = symbol.Split(':').Reverse().FirstOrDefault(); // accommodate Think or Swim group prefix
                    //if(!isNotesFileName.Success)
                    result = new AutoInitSortedDictionary<string, string>(); // In notes files we can have multiple lines for the same symbol so we use the same results
                    if(isSchwabGroupFileName.Success)
                    {
                        var groupOrManagedBy = isSchwabGroupFileName.Groups[3].Value;
                        var groupName = csv.GetField<string>(groupOrManagedBy);
                        result.Add(groupOrManagedBy, groupName);
                    }
                    else if (isAccountFileName.Success)
                    {
                        var columnName = isAccountFileName.Groups[3].Value;
                        var accountName = csv.GetField<string>(columnName).ToUpper();
                        result.Add("Account", accountName);
                    }
                    else if (isSchwabATR21DayFileName.Success)
                    {
                        var columnName = isSchwabATR21DayFileName.Groups[3].Value;
                        var atr21day = csv.GetField<string>(columnName);
                        result.Add(columnName, atr21day);
                    }
                    else if (isSchwabNetLiqPercentFileName.Success)
                    {
                        var columnName = isSchwabNetLiqPercentFileName.Groups[3].Value;
                        var netLiqPercent = csv.GetField<string>(columnName);
                        result.Add(columnName, netLiqPercent);
                    }
                    else if (isSymbolFileName.Success)
                    {
                        var symbolName = csv.GetField<string>("SymbolName");
                        result.Add("SymbolName", symbolName);
                    }
                    else if (isDecoratedSymbolFileName.Success)
                    {
                        // @@todo@@ The resulting spreadsheet cells show up as purple when this file ("sss Schwab
                        // DecoratedSymbol.csv") has been updated (on the weekend, for example) by
                        // "ExtractTOSStockGroups.py" (which is called by the bash shell script to use this and
                        // other programs to create the excel spreadsheet) and the table "defaultData" has not
                        // been updated on account of the master data files (MinDollarVol10MComp50.csv &
                        // MinDollarVol10MComp50ETF.csv) have not been updated with a fresh download from Market
                        // Surge. You typically don't want to update these on the weekend without updating all the
                        // data files.

                        // The solution is start with the date of this file and count backwards, day by day until
                        // we find an entry in "defaultValues" and make a new entry in "defaultData" and copy the
                        // latest data to this new entry.

                        // This ("sss Schwab DecoratedSymbol.csv") is a file that contains the decorated symbol
                        // name, which is an equity symbol prefixed with (typically) a zero filled integer that
                        // implements the sort order.

                        // @@todo@@: We are assuming that the master data files (in the map defaultValues which is
                        // loaded from the files MinDollarVol10MComp50.csv & MinDollarVol10MComp50ETF.csv) were
                        // updated the last time the market was open! This might not be true! If this is Sunday
                        // they may have been updated on Saturday or if Friday was a market holiday, they could be
                        // as old Thursday.

                        DateTime ldt = AdjustDateForMarketHolidays(currentDataFileDateTime);

                        // @@todo@@ accommodate other market holidays.
                        if (defaultValues is not null)
                        {
                            if (defaultValues.ContainsKey(symbolOnly))
                            {
                                if (defaultValues[symbolOnly].ContainsKey(ldt.Date))
                                {
                                    // copy the most recent data in the defaultValues table to a new entry in defaultValues table that matches the current
                                    // file's date so we can find it later.
                                    foreach (var (key, val) in defaultValues[symbolOnly][ldt.Date])
                                    {
                                        if (!result.ContainsKey(key))
                                        {
                                            result.Add(key, val);
                                            SaveMinMaxDoubleValues(key, val, symbolOnly);
                                        }
                                    }
                                }
                                else if (!generateMasterRegexList)
                                {
                                    WriteLine($"Error: DecoratedSymbol file \"{fileName}\" has no data for {symbolOnly}");
                                }
                            }
                            else if (!generateMasterRegexList)
                            {
                                WriteLine($"Error: table DefaultValues has no data for {symbolOnly} ");
                            }
                        }
                        else if (!generateMasterRegexList)
                        {
                            // This is normal when generating the defaultValues
                            //WriteLine($"Error: table DefaultValues is null for {symbolOnly} ");
                        }
                        var decoratedSymbolName = csv.GetField<string>("DecoratedSymbol");
                        result.Add("DecoratedSymbol", decoratedSymbolName);
                    }
                    if (isUnrealizedGainName.Success)
                    {
                        name = isUnrealizedGainName.Groups[3].Value;
                        var sharesOrProfitLossOrUnitCost = csv.GetField<string>(name);
                        result.Add(name, sharesOrProfitLossOrUnitCost);
                    }
                    else if (isSymbolFileName.Success)
                    {
                        name = isSymbolFileName.Groups[2].Value;
                        result.Add(name, symbol);
                    }
                    else if (isNotesFileName.Success)
                    {
                        var noteDaysOld = 0.0;
                        name = isNotesFileName.Groups[2].Value;
                        var isFormulas = isNotesFileName.Success ? isNotesFileName.Groups[3].Value == "Formulas" : false;
                        var dateString = csv.GetField<string>("Date");
                        if (DateTime.TryParse(dateString, out DateTime date))
                        {
                            noteDaysOld = (TODAY - date).TotalDays;
                            dateString = date.ToString("ddd MMM dd, yy");
                        }
                        var note = csv.GetField<string>(isFormulas?"Formulas":"Notes"); // note can contain comment, formulas, or note
                        note = convertPunctuationsToXML(note, patRestoreCommas, ",");
                        note = convertPunctuationsToXML(note, patRestoreNewLines, "&#10;");
                        if (!baseNameNoVersionFn.Contains("TRADE"))
                            foreach ((var pat, var replace) in patHighLightKeyWord)
                                /*if (pat.Match(note).Success)*/ note = pat.Replace(note, replace);
                        if (noteDaysOld > 99) noteDaysOld = 99;
                        if (!string.IsNullOrWhiteSpace(note))
                        {
                            HashSet<string> tags;
                            (tags,note) = ExtractTags(note);
                            notes = ((isCommentsFileName | isFormulasFileName) ? "" : $"{noteDaysOld.ToString("00")} {dateString} <B>{symbol}</B>: {string.Join(" " , tags)}") + note;// leave that date out of comments so they will take up less room
                        }
                        // Separate multiple notes with a new line
                        //var newline = result.ContainsKey($"{name}Notes") ? "&#10;" : "";
                        if (result.ContainsKey($"{name}Notes"))
                        {
                            if (noteDaysOld < maxNotesDaysOld) // @@todo@@ Does this ever execute? I think NOT! Remove this! I think this is now superfluous since we get a new result every time thru
                            {
                                var oldSize = result[$"{name}Notes"].Length;
                                var newAdditional = note.Length;
                                if (oldSize+newAdditional < maxNoteSize)
                                {
                                    // before joining the new note with the old notes, we need to extract the tags from the new note and the old notes (notes) and put them at the head (of the new note)
                                    var newNote = result[$"{name}Notes"];
                                    HashSet<string> newTags;
                                    (newTags, newNote) = ExtractTags(newNote);
                                    HashSet<string> oldTags;
                                    (oldTags, notes) = ExtractTags(notes);

                                    var m = patNoteHeaderBody.Match(newNote);
                                    if (m.Success)
                                    {
                                        // replace group 4 with union of old and new tags
                                        var newTagsString = string.Join(" ", oldTags.Union(newTags));
                                        newNote = m.Groups[1].Value + " " + newTagsString;
                                    }

                                    result[$"{name}Notes"] = newNote + "&#10;" + notes;
                                    var newSize = result[$"{name}Notes"].Length;
                                    if (debug1) WriteLine($"""Adding note:      {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld}) oldSize={oldSize} adding={newAdditional} newSize={newSize}""");
                                }
                                else
                                {
                                    if (debug1) WriteLine($"""note too big:     {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld}) oldSize={oldSize} adding={newAdditional}""");
                                    skip = true;
                                }
                            }
                            else
                            {
                                if (debug1) WriteLine($"""Too old:          {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld})""");
                                skip = true;
                            }
                        }
                        else
                        {
                            if (noteDaysOld < maxNotesDaysOld)
                            {
                                if (debug1) WriteLine($"""Initial Add:      {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld})""");
                                result.Add($"{name}Notes", notes);
                            }
                            else
                            {
                                if (debug1) WriteLine($"""Initial Add skipBaseFile: {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld})""");
                                skip = true;
                            }
                        }
                    }
                    else // this is a stock list (as opposed to a profit loss or notes file)
                    {
                        //var Symbol = csv.GetField<string>("Symbol");
                        //WriteLine($"---------- Begin {Symbol} in {fileName}-------------------");
                        var eventDate = csv.GetField<string>("Event Date");
                        var isNullEventDate = string.IsNullOrEmpty(eventDate);
                        // Some stock files, like "RS Line Blue Dot" have event data and we are only interested in the more recent events.
                        if (isNullEventDate || ValidEventDate(currentDataFileDateTime, eventDate, maxEventAge)) // Events for today only, skipBaseFile old events
                        {
                            if (defaultValues != null)
                            {                                
                                if (defaultValues.ContainsKey(symbolOnly))
                                {
                                    if (defaultValues[symbolOnly] != null)
                                    {
                                        var defaultValuesForSymbol = defaultValues[symbolOnly];

                                        // If the market is open and has not closed yet (it is before 1pm PST), use default values from the
                                        // previous day if they are not available for the current day.

                                        // @@todo@@ If this is a weekend day or a market holiday, count backwards day by day until we have a
                                        // non-market holiday and copy that data for the non-market holiday to a newly created entry in
                                        // defaultData that matches the date for this data file.

                                        // assume the most recent data in the defaultValues table matches the date of the file.
                                        var defaultDataDateTime = defaultValuesForSymbol.Keys.Max();// @@todo@@ use the date time of the master list files instead.
                                        var ldt = currentDataFileDateTime; // simplify logic and perform deep copy and save the original value just in case we need it some day.
                                        var weHaveDefaultDataForToday = defaultValuesForSymbol.ContainsKey(ldt.Date);
                                        var weHaveDefaultData = weHaveDefaultDataForToday;
                                        var weHaveDefaultDataForYesterday = defaultValuesForSymbol.ContainsKey(currentDataFileDateTime.AddHours(-allowOldDataFilesHoursOld).Date);
                                        if(weHaveDefaultDataForToday)
                                        {
                                            // we have data for today, use it.
                                            ldt = currentDataFileDateTime;
                                            weHaveDefaultData = true;
                                            if (debug1 && !generateMasterRegexList)
                                                WriteLine($"We have default data for {symbol} in {fileName} in the master list");
                                        }
                                        else if (weHaveDefaultDataForYesterday)
                                        {
                                            ldt = currentDataFileDateTime.AddHours(-allowOldDataFilesHoursOld);
                                            if(debug1 && !generateMasterRegexList)
                                                WriteLine($"We have default data for yesterday: defaultValuesForSymbol.ContainsKey(ldt.Date)={defaultValuesForSymbol.ContainsKey(ldt.Date)};");
                                            weHaveDefaultData = true;

                                        }
                                        else if (currentDataFileDateTime.Date <= defaultDataDateTime.Date)
                                        {
                                            // we are processing a data file that is older than the most recent data in the defaultValues table
                                            // this is OK: we will just use the more recent data
                                            ldt = defaultDataDateTime;
                                            if (debug1 && !generateMasterRegexList)
                                                WriteLine($"Using most recent default data for {symbolOnly} in {fileName} for {currentDataFileDateTime.Date:yyyy-MM-dd} because we are processing an old data file for date {currentDataFileDateTime:yyyy-MM-dd}");
                                            weHaveDefaultData = true;
                                        }
                                        else
                                        {
                                            // We are processing a data file that is newer than the most recent data in the defaultValues table
                                            // this is NOT OK: we will need to copy the most recent data to a new entry in the defaultValues table
                                            // with the date of the current data file so we can find it later.
                                            ldt = defaultDataDateTime;
                                            if (debug1 && !generateMasterRegexList)
                                                WriteLine($"Info: Is today a market holiday? default data for {symbolOnly} for {currentDataFileDateTime:yyyy-MM-dd} in {fileName} is more recent than {defaultDataDateTime.Date:yyyy-MM-dd} for default data (master list) so don't copy");
                                            weHaveDefaultData = false; // @@todo@@: this is not working, why not?
                                            weHaveDefaultData = true;
                                        }
                                        if (weHaveDefaultData)
                                        {
                                            if (defaultValues[symbolOnly][ldt.Date] != null)
                                            {
                                                // copy the current data from the default values
                                                foreach (var (key, val) in defaultValues[symbolOnly][ldt.Date])
                                                {
                                                    if (!result.ContainsKey(key))
                                                    {
                                                        result.Add(key, val);
                                                        SaveMinMaxDoubleValues(key, val, symbolOnly);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            // Is this necessary? Can we just get everything from the defaultvalues? I think so. Unless we are loading the default values!
                            foreach (var (_, attrNameAndType) in mapPositionToAttributeName)
                            {
                                if (attrNameAndType.name != "Days to Earnings")
                                {
                                    var attrValue = csv.GetField<string>(attrNameAndType.name);
                                    if (string.IsNullOrEmpty(attrValue) && defaultValues != null && defaultValues.ContainsKey(symbol) && defaultValues[symbol] != null && defaultValues[symbol].ContainsKey(currentDataFileDateTime) && defaultValues[symbol][currentDataFileDateTime] != null && defaultValues[symbol][currentDataFileDateTime].ContainsKey(attrNameAndType.name))
                                    {
                                        attrValue = defaultValues[symbol][currentDataFileDateTime][attrNameAndType.name];
                                        //WriteLine($"fileName={fileName} symbol={symbol} name={attrNameAndType.name} value={attrValue}");
                                    }
                                    if (attrNameAndType.name == "EPS Due Date")
                                    {
                                        if (string.IsNullOrEmpty(attrValue) || attrValue == "-")
                                        {
                                            if (!result.ContainsKey("Days to Earnings"))
                                                result.Add("Days to Earnings", "999");
                                            if (!result.ContainsKey(attrNameAndType.name))
                                                result.Add(attrNameAndType.name, attrValue);
                                        }
                                        else
                                        {
                                            var earningsDate = DateTime.Parse(attrValue);
                                            var daysUntilEarnings = (earningsDate - DateTime.Now).TotalDays + 1;
                                            if (daysUntilEarnings < 0)
                                            {
                                                daysUntilEarnings = 365+daysUntilEarnings;
                                            }
                                            if (!result.ContainsKey("Days to Earnings"))
                                                result.Add("Days to Earnings", ((long)(daysUntilEarnings + 0.5)).ToString());
                                            attrValue = earningsDate.ToString("yyyy-MM-dd ddd");
                                            if (!result.ContainsKey(attrNameAndType.name))
                                                result.Add(attrNameAndType.name, attrValue);
                                            if (!result.ContainsKey("seq"))
                                                result.Add("seq", (count - 1).ToString("00")); // this is for lists like IBD 50 where the order is important
                                            if (!result.ContainsKey(attrNameAndType.displayName))
                                                result.Add(attrNameAndType.displayName, attrNameAndType.convert(attrValue));
                                        }
                                    }
                                    else if (!result.ContainsKey(attrNameAndType.name))
                                    {
                                        //result.Add(attrNameAndType.displayName, attrNameAndType.convert(attrValue)); // @@bug@@ is this superfluous?
                                        result.Add(attrNameAndType.name, attrValue);
                                        SaveMinMaxDoubleValues(attrNameAndType.name, attrValue, symbol);
                                        //WriteLine($" add key: \"{attrNameAndType.name}\" val: \"{attrValue}\" filename={fileName}");
                                        /*
                                        if (result.ContainsKey(attrNameAndType.name))
                                        {
                                            WriteLine($" contains key: \"{attrNameAndType.name}\" val: \"{result[attrNameAndType.name]}\" filename={fileName}");
                                        }
                                        else
                                        {
                                            WriteLine($" add key: \"{attrNameAndType.name}\" val: \"{attrValue}\" filename={fileName}");
                                            result.Add(attrNameAndType.name, attrValue);
                                        }
                                        */
                                    }
                                }
                            }
                            //if(!result.ContainsKey("seq"))
                            result["seq"] = (count - 1).ToString("00"); // this is for lists like IBD 50 where the order is important
                            if (!result.ContainsKey("Updated"))
                                result.Add("Updated", new FileInfo(fileName).LastWriteTime.ToString("ddd MMM dd yy"));
                        }
                        else
                        {
                            //WriteLine($"skipping eventDate={eventDate}");
                        }
                    }
                    var symbols = /*symbol==symbolOnly*/ true ? new[] { symbol } : new[] { symbol, symbolOnly };
                    foreach (var sym in symbols)
                    {
                        if(string.IsNullOrEmpty(sym))  continue;                        
                        if (skip)
                        {
                            if (debug1) WriteLine($"adding row SKIP    [{sym}][{currentDataFileDateTime}] from file={fileName}");
                        }
                        else if (rows.ContainsKey(sym) && rows[sym].ContainsKey(currentDataFileDateTime))
                        {
                            foreach ((var key, var val) in result)
                            {
                                var oldNotes = val;
                                // concatenate the new value with the old value: @@todo@@: implement constraints on size and date of hold notes here!
                                var newNote = rows[sym][currentDataFileDateTime][key];
                                // if newNote is not null, then we need to extract the tags from the new note and the old notes (notes) and put them at the head (of the new note)
                                if (newNote is not null) 
                                {
                                    HashSet<string> newTags;
                                    (newTags, newNote) = ExtractTags(newNote);
                                    HashSet<string> oldTags;
                                    (oldTags, oldNotes) = ExtractTags(oldNotes);

                                    //patNoteHeader = new Regex(@"^(\d{2}\s+(Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d+,?\s+\d{2}\s*<B>[A-Z]+</B>:)(\s*(\s*![A-Z]+!\s*))*(.*)$");
                                    var m = patNoteHeaderBody.Match(newNote);
                                    if (m.Success)
                                    {
                                        // replace group 4 with union of old and new tags
                                        var newTagsString = string.Join(" ", oldTags.Union(newTags));
                                        newNote = m.Groups[1].Value + " " + newTagsString + " " + m.Groups[5].Value.Trim() + " " + m.Groups[6].Value.Trim();
                                    }
                                    rows[sym][currentDataFileDateTime][key] = newNote + "&#10;" + oldNotes;
                                }
                            }
                        }
                        else
                        {
                            try
                            {
                                rows[sym][currentDataFileDateTime] = result;
                                if (debug1) WriteLine($"adding row         [{sym}][{currentDataFileDateTime}] from file={fileName}");
                            }
                            catch (System.ArgumentException)
                            {
                                rows[sym][currentDataFileDateTime][$"{name}Notes"] = notes + "&#10;" + rows[sym][currentDataFileDateTime][$"{name}Notes"];
                                if(debug1) WriteLine($"appending row         [{sym}][{currentDataFileDateTime}][{name}Notes] notes={notes} from file={fileName}");
                            }
                        }
                    }
                }
                else
                {
                    var more = csv.ReadHeader();
                }
            }
        }
        return rows;

        static bool ValidEventDate(DateTime listDateTime, string eventDate, int maxEventAge)
        {
            if (maxEventAge == int.MaxValue)
            {
                return true;
            }
            else
            {
                var eventDateMatch = patDate_YYYY_MM_dd.Match(eventDate);
                var eventDateIsValidDate = eventDateMatch.Success;
                if (eventDateIsValidDate)
                {
                    var eventAge = (listDateTime.Date - new DateTime(Int32.Parse(eventDateMatch.Groups[1].Value), Int32.Parse(eventDateMatch.Groups[2].Value), Int32.Parse(eventDateMatch.Groups[3].Value))).TotalDays;
                    var result = eventDateMatch.Success && (eventAge <= maxEventAge);
                    return result;
                }
                else
                {
                    return false;
                }
            }
        }
    }
    /// <summary>
    /// Adjusts a given date to account for market holidays by shifting it to the previous Friday if it falls on a
    /// weekend.
    /// @@todo@@ Accommodate other holidays besides Saturday and Sunday.
    /// @@todo@@ Accommodate the case when this is a Sunday and the master data files in defaultValues has been updated on Saturday.
    /// </summary>
    /// <param name="listDateTime">Represents the date and time of the CSV data file being processed.</param>
    /// <returns>Returns the adjusted date as a DateTime object.</returns>
    static DateTime AdjustDateForMarketHolidays(DateTime listDateTime)
    {
        var ldt = listDateTime; // date time of the csv data file.
        if (listDateTime.DayOfWeek == DayOfWeek.Sunday)
            ldt = ldt.AddDays(-2); // if today is Sunday, then the data file is from Friday
        else if (listDateTime.DayOfWeek == DayOfWeek.Saturday)
            ldt = ldt.AddDays(-1); // if today is Saturday, then the data file is from Friday
        return ldt;
    }

    private static (HashSet<string>, string) ExtractTags(string note)
    {
        var tags = new HashSet<string>();
        // remove all the tags from notes and add them to the tags
        var m = patTag.Match(note);
        while (m.Success)
        {
         var tag = m.Groups[0].Value;
            note = note.Replace(tag, "");            
            tags.Add(tag);
            m = m.NextMatch();
        }

        return (tags,note);
    }

    void SaveMinMaxDoubleValues(string key, string val, string symbol)
    {
        if(!string.IsNullOrEmpty(val) && val != "-" && mapAttributeNametoAttributeAttributes.ContainsKey(key) && mapAttributeNametoAttributeAttributes[key].numeric)
        {
            if (doubleMin.ContainsKey(key))
            {
                var valDouble = double.Parse(val);
                //if (valDouble > 50 && key == "Price % Chg") WriteLine($"symbol = {symbol} key={key} val={val}");
                if (valDouble < doubleMin[key])
                    doubleMin[key] = valDouble;
            }
            else
            {
                doubleMin[key] = double.Parse(val);
            }
            if (doubleMax.ContainsKey(key))
            {
                var valDouble = double.Parse(val);
                //if (valDouble > 50 && key == "Price % Chg")  WriteLine($"symbol = {symbol} key={key} val={val}");
                if (valDouble > doubleMax[key])
                    doubleMax[key] = valDouble;
            }
            else
            {
                doubleMax[key] = double.Parse(val);
            }
        }
    }

    private static string convertPunctuationsToXML(string? note, Regex patOld, string newString)
    {
        var matchCollection = patOld.Matches(note); foreach (var match in matchCollection) note = note.Replace(match.ToString(), newString);
        return note;
    }

    // https://stackoverflow.com/questions/381366/is-there-a-wildcard-expansion-option-for-net-apps
    static List<string> ExpandFilePaths(string[] args)
    {
        var fileList = new List<string>();

        foreach (var arg in args)
        {
            var substitutedArg = ExpandEnvironmentVariables(arg);

            var dirPart = Path.GetDirectoryName(substitutedArg);
            if (dirPart.Length == 0)
                dirPart = ".";

            var filePart = Path.GetFileName(substitutedArg);

            foreach (var filepath in Directory.GetFiles(dirPart, filePart))
                fileList.Add(filepath);
        }

        return fileList;
    }

    static Regex patContractEnvironmentVars = new Regex(@"^([cC]:\\Users\\shein\\Downloads\\)(.*)$");
    static bool debug1 = false;

    static SortedSet<string> GetAllCsvFilesFromDownloadDirectory()
    {
        var fileNames = new SortedSet<string>();
        foreach (var fileName in Directory.GetFiles(ExpandEnvironmentVariables(@"%USERPROFILE%\Downloads"), "*.csv"))
        {
            if (!SkipFile(fileName))
            {
                fileNames.Add(fileName);
            }
        }
        return fileNames;
    }

    static string[] GetFileNamesAndSwitchesFromArgs(string[] args)
    {
        var fns = args.ToList();
        var noWildCards = false;
        foreach (var fn in fns)
        {
            if (fn == "--noWildCards")
                noWildCards = true;
        }
        if (!noWildCards)
            fns.AddRange(GetAllCsvFilesFromDownloadDirectory());
        return fns.ToArray();
    }

}
