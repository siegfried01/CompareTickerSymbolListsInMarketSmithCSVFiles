using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using CompareTickerSymbolListsInCSVFiles;
using CsvHelper;
using static System.Console;
using static System.Environment;
using static System.Net.WebRequestMethods;
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
record DataFileCatagories(SortedSet<string> preferred, SortedSet<string> missingDataColums, SortedSet<string> donotknow);
internal class CompareTickerSymbolListsInCSVFilesMainProgram
{
    static SortedSet<string> filesWithMissingDataColumns = new SortedSet<string>(){
        @"All RS Line New High.csv",
        @"Deletions.csv",
        @"Established Downtrend.csv",
        @"Extended Stocks.csv",
        @"IBD 50 Index.csv",
        @"IBD Big Cap 20.csv",
        @"IBD Live Ready.csv",
        @"RS Line New High.csv",
        @"RS Line Blue Dot.csv",
        @"IBD Live Watch.csv",
    };
    static SortedSet<string> preferredFiles = new SortedSet<string>() {@"Additions.csv",
        @"LB Sector Leaders.csv",
        @"LB Watch.csv",
        @"Long Term Leaders Watch.csv",
        @"Long Term Leaders.csv",
        @"ttt ML Holdings NetLiquidValue.csv",
        @"ttt ML Holdings PortionOfTotalAccount.csv",
        @"ttt ML Holdings Price.csv",
        @"ttt ML Holdings ProfitLoss.csv",
        @"ttt ML Holdings ProfitLossDollar.csv",
        @"ttt ML Holdings Shares.csv",
        @"ttt ML Holdings UnitCost.csv",
        @"ttt ML Holdings.csv",
        @"ttt Swad Jul 17 Buy.csv",
        @"ttt Swad Jul 17 Uni.csv",
        @"Webby's Daily KISS.csv",
        @"zzzBofANotes.csv",
        @"zzzCFRANotes.csv",
        @"zzzIBDLNNotes.csv",
        @"zzzMLDISCUSSNotes.csv",
        @"zzzMSTARNotes.csv",
        @"zzzNEWSMLNotes.csv",
        @"zzzNEWSMSNotes.csv" };

    static string orderPrefix = "aaa|bbb|ccc|ddd|eee|fff|ggg|hhh|iii|jjj|kkk|lll|mmm|nnn|ooo|ppp|qqq|rrr|sss|ttt|uuu|vvv|www|xxx|yyy|zzz";
    static Regex patFileNameOrderPrefix = new Regex(@$"^(({orderPrefix})\s*)?(.*)$"); // optional file name prefix for implementing the order of the columns
    static Regex patFileExtension = new Regex("^([^\\.]+)\\.([^\\.]+)$");
    static Regex patGotoRow = new Regex("^--GotoRow=([0-9]+)$");
    //[GeneratedRegex(@"zzz([a-zA-Z0-9]*)Notes(\.csv)?$")]
    static Regex patNotes = new Regex(@$"({orderPrefix})([a-zA-Z0-9]*)\s*Notes(\.csv)?$");
    static Regex patShares = new Regex(@$"({orderPrefix}) ?([- a-zA-Z0-9]*)\s*Shares(\.csv)?$");
    static Regex patUnrealizedGains = new Regex(@$"({orderPrefix}) ?([- a-zA-Z0-9]*)\s*(Price|Shares|UnitCost|ProfitLoss(Dollar)?|NetLiquidValue|PortionOf(Total)?Account)(\.csv)?$");
    static Regex patHistory = new Regex(@"--[Hh]istoryDays=([0-9]+)");
    static Regex patAddSpaceNotes = new Regex(@"^(.*)(Notes)$");
    static Regex patNoteColumnWidth = new Regex(@"--Note(Col(umn)?Width)?=([0-9]+)");
    static Regex patMaxHistory = new Regex(@"--MaxHistory=([0-9]+)");
    static Regex patMaxNotesDaysOld = new Regex(@"--MaxNotesDaysOld=([0-9]+(\.[0-9]+))");
    static Regex patTickerSymbolColumnWidth = new Regex(@"--(Ticker)?Symbol(Col(umn)?Width)?=([0-9]+)");
    static string NoteColumnWidth = "400";
    static string SymbolColumnWidth = "37";//"32.75";
    static List<Regex> patSkipFiles = new List<Regex> {
        new Regex(@"^MinDollarVol10MComp50\.csv$"),
        new Regex (@"^197 Industry Groups\.csv$"),
        new Regex(@"Merrill-Holdings-Unrealized-Gain-Loss-Summary\.csv$"),
        new Regex(@"^Merrill.*(Holdings?|All).*\.csv$"),
        new Regex(@"^ExportData.*\.csv$"),
        new Regex(@"PositionStatement\.csv$"),
        new Regex(@"^Realized Gain Loss.*\.csv$"),
        new Regex(@"^SEP-IRA-Positions-\d+-\d+-\d+-\d+\.csv$"),
        new Regex(@"^paper trading realized gain loss TOS \d+-\d+-\d+-AccountStatement\.csv$"),
        new Regex(@"^(Merril.*|SEP-IRA.*|paper trading.*|.*Ameritrade.*|197 Industry Groups|MinDollarVol[0-9]+MComp[0-9]+)\.csv$")};
    static List<Regex> patSkipFilesMaterList = new List<Regex> { new Regex(@"^zzz.*\.csv$") };
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
            switchValue = string.Empty;
            return false;
        }
    }
    static double maxNotesDaysOld = 30;
    static int maxNoteSize = 1000;
    public static void Main(string[] args)
    {
        int BACK_HISTORY_COUNT = 60;
        string initialRow = args.Length > 0 && Int32.TryParse(args[0], out Int32 _) ? args[0] : "6";
        var generateCSV = false;
        var generateXML = true;
        var generateMasterList = false;
        AutoMultiDimSortedDictionary<string/*file name*/, AutoMultiDimSortedDictionary<string /*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*metric name*/, string/*metric value*/>>>> mapFileNameToSymbolToDatesToAttributes = new();// new DescendingComparer<string>());
        SortedDictionary<string, SymbolInList> comparisonGrid = new();
        SortedDictionary<string, int> mapFileNameToColumnPosition = new(new DescendingComparer<string>());
        SortedDictionary<string, string> mapFileNameToFileNameNoExt = new();
        AutoMultiDimSortedDictionary<string/*symbol*/, AutoInitSortedDictionary<string/*file name only*/, SortedSet<DateTime>>> mapSymbolToFileNameToDates = new();
        SortedDictionary<string, DateTime> mapFileNameToMostRecentFileDate = new();
        SortedDictionary<DateTime, string> mapMostRecentDateToFile = new(new DescendingComparer<DateTime>());
        var sbXMLRows = new StringBuilder();
        var sbRowHeadersRow = new StringBuilder();
        var sbStyles = new StringBuilder();
        var sbCSV = new StringBuilder();
        var masterList = new List<string>();
        var workSheetName = "CompareSymbols";
        var debug = false;
        var rowCount = 0;
        var daysOfHistory = 5;
        var columnCurrent = 0;
        var argsAndDownloadFiles = GetFileNamesAndSwitchesFromArgs(args);
        var xmlColumnWidths = new List<string> { """<Column ss:Width="15" />""" }; // why does the width have no affect?
        var verticalSplitter = false;
        var excelStyles = new AutoInitSortedDictionary<string/* Hexadecimal color name */, ExcelStyle>();
        // emacs solve([0 * m + b = 120, 1 * m + b = 235], [m, b])==[m = 115, b = 120]
        var stockExcelHueAgeStyle = new ExcelStyleHueRange { Saturation = 240f, Luminance = 240f, HueMin = 120.0 / 360.0 * 255.0, HueMax = 235 / 360.0 * 255.0, InputMetricMin = 0, InputMetricMax = 1, HueScale = 115.0, HueOffset = 120 };
        var stockExcelSaturationAgeStyle = new ExcelStyleSaturationRange { Hue = 120.0 / 360.0 * 255.0, Luminance = 240f, SaturationMin = 120.0 / 360.0 * 255.0, SaturationMax = 355.0 / 360.0 * 255.0, InputMetricMin = 0, InputMetricMax = 1, SaturationScale = 115.0, SaturationOffset = 120 };

        IEnumerable<History> history = null;
        SortedSet<DateTime> historyDates = new(new DescendingComparer<DateTime>());
        SortedSet<string> fileNamesWithHistory = new();

        foreach((var position, (var name, var numeric)) in mapPositionToAttributeName)
        {
            mapAttributeNameToPosition.Add(name, position);
        }

        foreach (string arg in argsAndDownloadFiles/*args*/)
        {
            var matchHistory = patHistory.Match(arg);
            if (arg == "--debug")
            {
                debug = true;
            }
            else if (matchHistory.Success)
            {
                daysOfHistory = Int32.Parse(matchHistory.Groups[1].Value);
            }
            else if (arg == "--CSV")
            {
                generateCSV = true;
                generateXML = false;
            }
            else if (arg == "--Reverse")
            {
                DescendingComparer<string>.reverse = true;
                verticalSplitter = true;
            }
            else if (arg == "--VerticalSplitter")
                verticalSplitter = true;
            else if (arg == "--noWildCards")
            {
                // skip, already handled
            }
            else if (arg == "--masterList")
            {
                generateMasterList = true;
                patSkipFiles.AddRange(patSkipFilesMaterList);
                generateXML = false;
                generateCSV = false;
            }
            else if (MatchSwitchWithValue(patNoteColumnWidth, arg, out string width, 3))
            {
                NoteColumnWidth = width;
            }
            else if (MatchSwitchWithValue(patGotoRow, arg, out string row))
            {
                initialRow = row;
            }
            else if (MatchSwitchWithValue(patMaxNotesDaysOld, arg, out string daysOld))
            {
                maxNotesDaysOld=double.Parse(daysOld);
            }
            else
            {
                //var fileNames = ExpandFilePaths(arg);
                var fileName = ExpandEnvironmentVariables(arg);
                var skip = SkipFile(fileName);

                if (System.IO.File.Exists(fileName) && !skip)
                {
                    try
                    {
                        //DateTime lastWriteTime = System.IO.File.GetLastWriteTime(fileName);
                        var listDateTime = ExtractDateTimeFromFilePath(fileName);
                        AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*metric name*/, string/*metric value*/>>> csvCurrentData = ParseCSV(fileName, listDateTime);
                        fileName = Path.GetFileName(fileName);
                        if(generateMasterList)
                            history = new List<History>();
                        else
                            history = Directory.GetDirectories(ExpandEnvironmentVariables(@"%USERPROFILE%\Downloads")).Select(d => new { Name = Path.GetFileName(d), Path = d }).Where(d => patDateTime_YYYY_MMM_dd_ddd.Match(d.Name).Success).Select(d => { var m = patDateTime_YYYY_MMM_dd_ddd.Match(d.Name); var r = new { Name = d.Name, Path = d.Path, Date = System.DateTime.Parse(m.Groups[1].Value) }; return r; }).OrderByDescending(d => d.Date).Take(BACK_HISTORY_COUNT).Select(f => new History { Date = f.Date, FileExists = System.IO.File.Exists(Path.Combine(f.Path + "\\" + fileName)), Name = System.IO.Path.Combine(f.Path + "\\" + fileName) });
                        mapMostRecentDateToFile[listDateTime.Date] = fileName;
                        mapFileNameToMostRecentFileDate[fileName] = listDateTime.Date;
                        MakeHistoryDateSet(history, TODAY, ref historyDates);
                        UpdateMapSymbolToFileNameToDates(mapSymbolToFileNameToDates, fileName, csvCurrentData);
                        UpdateMapFileNameToSymbolToDateToAttributes(mapFileNameToSymbolToDatesToAttributes, fileName, csvCurrentData);
                        if (!generateMasterList)
                        {
                            foreach (var oldFile in history)
                            {
                                if (oldFile.FileExists)
                                {
                                    var csvHistoryData = ParseCSV(oldFile.Name, oldFile.Date);
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
                        mapFileNameToColumnPosition[fileName] = 0; // All filenames are in column zero for now, fix this later
                        columnCurrent++;
                    }
                    catch (IOException ex)
                    {
                        WriteLine($"skipping file {fileName} Error={ex}");
                    }
                }
                else if (!generateMasterList)
                {
                    WriteLine(((skip ? "Skip file" : "File not found") + ": ") + fileName);
                }
            }
        }
        foreach ((var fileName, var mapSymbolToDatesToMetrics) in mapFileNameToSymbolToDatesToAttributes)
        {
            foreach ((var symbol, var mapDatesToMetrics) in mapSymbolToDatesToMetrics)
            {
                if (debug) WriteLine($"file={fileName} Adding {symbol} to comparisonSheet for {string.Join(",", mapDatesToMetrics.Keys.OrderByDescending(x => x).Select(x => x.Date.ToString("yy-MMM-dd-ddd")).Take(5))}");
                var mostRecentDateForThisFile = mapDatesToMetrics.Keys.Max(x => x.Date);
                if (mostRecentDateForThisFile == mapFileNameToMostRecentFileDate[fileName].Date)
                {
                    if (comparisonGrid.ContainsKey(symbol))
                    {
                        comparisonGrid[symbol].Lists.Add(fileName);
                    }
                    else
                    {
                        var tmp = new SortedSet<string>();
                        tmp.Add(fileName);
                        comparisonGrid.Add(symbol, new SymbolInList { Name = symbol, Lists = tmp });
                    }
                }
            }
        }
        var columnCount = 0;
        var fileNames = new string[mapFileNameToColumnPosition.Count];
        foreach (var fn in mapFileNameToColumnPosition.Keys) fileNames[columnCount++] = fn; // deep copy
        columnCount = 0;
        foreach (var fileName in fileNames)
        {
            if (columnCount > 0 && generateCSV)
                sbCSV.Append(",");
            mapFileNameToColumnPosition[fileName] = columnCount++;
            var fileNameWithoutOptionalPrefix = patFileNameOrderPrefix.Match(fileName).Groups[3].Value;
            var match = patFileExtension.Match(fileNameWithoutOptionalPrefix);
            var fileNameWithOutExtension = match.Success ? match.Groups[1].Value : Path.GetFileName(fileName);
            if (generateCSV)
            {
                sbCSV.Append($"\"{fileNameWithOutExtension}\"");
            }
            if (generateXML)
            {
                mapFileNameToFileNameNoExt.Add(fileName, fileNameWithOutExtension);
                var m = patNotes.Match(fileName);
                if (m.Success)
                {
                    xmlColumnWidths.Add($"""<Column ss:Width="{NoteColumnWidth}" />""");
                }
                else
                {
                    xmlColumnWidths.Add($"""<Column ss:Width="{SymbolColumnWidth}" />""");
                }
            }
        }
        if (generateCSV)
            sbCSV.AppendLine();
        else
            sbXMLRows.AppendLine();
        var historyDirectoryDateArray = historyDates.ToArray<DateTime>().Reverse<DateTime>().ToArray<DateTime>(); // why do I have to reverse this since it was created with a descending comparitor?
        foreach (var symbol in comparisonGrid.Keys.Where(s => !string.IsNullOrEmpty(s)))
        {
            columnCurrent = 0;
            if (generateMasterList)
            {
                masterList.Add(symbol);
            }
            if (generateXML)
            {
                sbXMLRows.AppendLine("""  <Row ss:AutoFitHeight="0">""");
                sbXMLRows.AppendLine($"""    <Cell><Data ss:Type="Number">{rowCount + 1}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>""");
            }
            bool skipOverBlankColumn = false;
            AutoMultiDimSortedDictionary<string/*fileName*/, AutoMultiDimSortedDictionary<string/*symbol*/, AutoInitSortedDictionary<string, string>>> latestAttributes = new();
            var latestFileNameForThisSymbol = string.Empty;
            var latestFileNamesSetsForThisSymbol = new DataFileCatagories(new SortedSet<string>(), new SortedSet<string>(), new SortedSet<string>());
            foreach (var fileName in DescendingComparer<string>.reverse ? comparisonGrid[symbol].Lists.Reverse() : comparisonGrid[symbol].Lists)
            {
                latestFileNameForThisSymbol = fileName;
                var index = mapFileNameToColumnPosition[fileName];
                if (debug) WriteLine($"fileName={fileName} index={index}");
                while (columnCurrent < index)
                {
                    if (/*columnCurrent > 0 &&*/ columnCurrent < columnCount)
                    {
                        if (generateCSV)
                        {
                            sbCSV.Append(",");
                        }
                        if (generateXML)
                        {
                            skipOverBlankColumn = true;
                            //sbXMLRows.AppendLine($"""    <Cell><Data ss:Type="String"></Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""");
                        }
                    }
                    columnCurrent++;
                }
                if (generateCSV && columnCurrent > 0 && columnCurrent < columnCount) // @@bug@@
                {
                    if (generateCSV)
                    {
                        sbCSV.Append(",");
                    }
                    if (generateXML)
                    {
                        var skipToIndex = skipOverBlankColumn ? $" ss:Index=\"{columnCurrent + 2}\"" : "";
                        skipOverBlankColumn = false;
                        sbXMLRows.AppendLine($"""    <Cell{skipToIndex}><Data ss:Type="String"></Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""");
                    }
                    columnCurrent++;
                }
                if (generateCSV)
                {
                    sbCSV.Append(symbol);
                }
                if (generateXML && columnCurrent >= 0 && columnCurrent < columnCount)
                {
                    var skipToIndex = skipOverBlankColumn ? $" ss:Index=\"{columnCurrent + 2}\"" : "";
                    skipOverBlankColumn = false;
                    var matchNotes = patNotes.Match(fileName);
                    var matchShares = patShares.Match(fileName);
                    var matchUnrealizedGains = patUnrealizedGains.Match(fileName);
                    var latest = mapFileNameToSymbolToDatesToAttributes[fileName][symbol].Keys.OrderByDescending(d => d).ToArray<DateTime>()[0];
                    var first = mapFileNameToSymbolToDatesToAttributes[fileName][symbol].Keys.FirstOrDefault();
                    if (matchNotes.Success)
                    {
                        var name = matchNotes.Groups[2].Value;
                        var stock = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                        var notes = stock[$"{name}Notes"];
                        notes = patDateTime_ddd_MMM_dd_YYYY_ddd_MMM_dd_YYYY.Replace(notes, m => m.Groups[6].Value); // remove the date of data entry and leave the date of the data (report)   
                        var line = $"""    <Cell{skipToIndex}><ss:Data ss:Type="String" xmlns="http://www.w3.org/TR/REC-html40">{notes}</ss:Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""";
                        sbXMLRows.AppendLine(line);
                    }
                    else if (matchUnrealizedGains.Success)
                    {
                        var name = matchUnrealizedGains.Groups[3].Value;
                        var stock = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                        var sharesOrProfitLossOrUnitCost = stock[name].Trim();
                        var style = "";
                        if ("ProfitLoss" == name || "PortionOfTotalAccount" == name)
                        {
                            var val = +(double.Parse(sharesOrProfitLossOrUnitCost) / 100);
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
                            var val = double.Parse(sharesOrProfitLossOrUnitCost);
                            if (val < 0)
                                style = " ss:StyleID=\"s74\"";
                            else
                                style = " ss:StyleID=\"s72\"";
                        }
                        var line = $"""    <Cell{skipToIndex}{style}><Data ss:Type="Number">{sharesOrProfitLossOrUnitCost}</ss:Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}-->""";
                        sbXMLRows.AppendLine(line);
                    }
                    else
                    {
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
                            var hd = historyDirectoryDateArray[jjj++].Date;
                            var dfts = datesForThisSymbolArray[kkk++].Date;
                            var eq = hd == dfts;
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
                                if (!datesForThisFileNameOnly.Contains(hd)) // we missed a downloading a file for this day. Assume the symbol was there.
                                {
                                    age++;
                                    kkk--; // back up and look for a match on the next go around.
                                }
                                else
                                {
                                    break;
                                }
                            }
                            count++;
                        }
                        latestAttributes[symbol][fileName] = mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest];
                        var attributes = string.Join(",", mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest].Select(entry => $"{entry.Key}={entry.Value}"));
                        attributes += $", days on list={age}";
                        var styleName = "s68";
                        if (age != 0 && fileNamesWithHistory.Contains(fileName))
                        {
                            var metric = 1.0 * age / (BACK_HISTORY_COUNT + 1);
                            metric = Math.Log(metric * (BACK_HISTORY_COUNT + 1)) / Math.Log(BACK_HISTORY_COUNT + 1);
                            stockExcelSaturationAgeStyle.InputMetric = metric;
                            var RGBHexColor = stockExcelSaturationAgeStyle.ColorHexRGB;
                            var style = new ExcelStyle { Color = RGBHexColor, Name = "s" + RGBHexColor };
                            excelStyles[RGBHexColor] = style;
                            styleName = "s" + RGBHexColor;
                            if (debug1) WriteLine($"f={fileName} symbol={symbol} age={age} metric={metric}");
                        }
                        var seqNoAndSymbol = fileName.Contains("IBD 50 Index") || fileName.Contains("LB Top 10") ? $"{mapFileNameToSymbolToDatesToAttributes[fileName][symbol][latest]["seq"]} {symbol}" : symbol;
                        sbXMLRows.AppendLine($"""    <Cell{skipToIndex} ss:StyleID="{styleName}" ss:HRef="https://marketsmith.investors.com/mstool?Symbol={symbol}&amp;Periodicity=Daily&amp;InstrumentType=Stock&amp;Source=sitemarketcondition&amp;AlertSubId=8241925&amp;ListId=0&amp;ParentId=0" x:HRefScreenTip="{attributes}"><Data ss:Type="String">{seqNoAndSymbol}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell><!-- col={columnCurrent} fn={fileName}--> """);
                    }
                }
                columnCurrent++;
            }
            // add other stock metrics here
            /**/

            rowCount++;
            if (generateXML){
                var styleName = "s68";
                columnCurrent = mapFileNameToFileNameNoExt.Keys.Count;
                var metricCells = ""; skipOverBlankColumn = true;
                var attributes = new AutoInitSortedDictionary<string, string>();
                var pFiles = latestAttributes[symbol].Keys.Intersect(preferredFiles);
                string bestFile = latestFileNameForThisSymbol;
                if (pFiles.Count() > 0)
                    bestFile = pFiles.FirstOrDefault();
                attributes = latestAttributes[symbol][bestFile];
                WriteLine($"Fetching latestAttributes for {symbol} from {bestFile}");
                foreach ((var position, (var attributeName, var numeric)) in mapPositionToAttributeName)//foreach ((var key, var val) in latestAttributes[symbol])
                {
                    var skipToIndex = skipOverBlankColumn ? $" ss:Index=\"{columnCurrent + 2}\"" : "";
                    var type = numeric ? "Number" : "String";
                    var value = attributes[attributeName];
                    metricCells += $"""    <Cell{skipToIndex} ss:StyleID="{styleName}"><Data ss:Type="{type}">{value}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>""";
                    columnCurrent++;
                    skipOverBlankColumn = false;
                }
                sbXMLRows.AppendLine(metricCells);
                sbXMLRows.AppendLine($"  </Row> <!-- {rowCount} -->");
            }
            if (generateCSV)
            {
                sbCSV.AppendLine();
            }
            if (generateXML)
            {
                sbXMLRows.AppendLine();
            }
        }
        var finalExcelOutputFilePath = ExpandEnvironmentVariables(@$"%USERPROFILE%\Downloads\CombineLists_{DateTime.Now.ToString("yyyy-MMM-dd-ddd-HH")}.{(generateCSV || generateMasterList ? "csv" : "xml")}");

        if (generateCSV && sbCSV.Length > 0)
        {
            System.IO.File.WriteAllText(finalExcelOutputFilePath, sbCSV.ToString());
            var excel = ExpandEnvironmentVariables(@"%MSOFFICE%\EXCEL.EXE");
            System.Diagnostics.Process.Start(excel, $"/s \"{finalExcelOutputFilePath}\"");
        }
        if (generateMasterList && masterList.Count() > 0)
        {
            WriteLine(string.Join("|", masterList)
                + "|" + masterList.Last() // bug fix for perl
                );
            /*
            File.WriteAllText(fn, masterList.ToString());
            var excel = ExpandEnvironmentVariables(@"%MSOFFICE%\EXCEL.EXE");
            System.Diagnostics.Process.Start(excel, $"/s \"{fn}\"");
            */
        }
        if (generateXML)
        {
            var headers = $"""<Cell ss:StyleID="s62"><Data ss:Type="String">Order</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>"""
                + $"""{string.Join("", (fileNames.OrderBy(fn => mapFileNameToColumnPosition[fn])).Select(fn => $"""<Cell ss:StyleID="s62"><Data ss:Type="String">{patAddSpaceNotes.Replace(mapFileNameToFileNameNoExt[fn], m => $"{m.Groups[1].Value} {m.Groups[2].Value}")}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>"""))} """
                + "";
            headers += $"""{string.Join("", (mapPositionToAttributeName.Select(positionNamePair => $"""<Cell ss:StyleID="s62"><Data ss:Type="String">{positionNamePair.Value.name}</Data><NamedCell ss:Name="_FilterDatabase"/></Cell>""")))}""";
            columnCount += mapPositionToAttributeName.Count;
            rowCount++;
            columnCount++; // include the extra column for the current row
            var stStyles = "";
            foreach ((var _, var v) in excelStyles)
            {
                stStyles += (string.IsNullOrEmpty(stStyles) ? "" : "\n    ") + v;
            }
            var worksheetXML = $"""
            <?xml version="1.0"?>
            <?mso-application progid="Excel.Sheet"?>
            <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
             xmlns:o="urn:schemas-microsoft-com:office:office"
             xmlns:x="urn:schemas-microsoft-com:office:excel"
             xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
             xmlns:html="http://www.w3.org/TR/REC-html40">
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
              {stStyles}
             </Styles>
             <Worksheet ss:Name="{workSheetName}">
               <Names>            
                <NamedRange ss:Name="_FilterDatabase" ss:RefersTo="={workSheetName}!R1C1:R{rowCount}C{columnCount}" ss:Hidden="1"/>
               </Names>
               <Table ss:ExpandedColumnCount="{columnCount}" ss:ExpandedRowCount="{rowCount}" x:FullColumns="1" x:FullRows="1" ss:DefaultColumnWidth="42" ss:DefaultRowHeight="11.25">
               {string.Join("", xmlColumnWidths)}
               <Row ss:Height="230.25">
                 {headers}
               </Row>
            """ + sbXMLRows.ToString();
            worksheetXML +=
             $""""
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
              """" + (verticalSplitter ?
              $""""
               <SplitHorizontal>4905</SplitHorizontal>
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
               <SplitHorizontal>4830</SplitHorizontal>
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
               <SplitHorizontal>4830</SplitHorizontal>
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
            </Workbook>            
            """";

            //WriteLine(worksheetXML);
            try
            {
                System.IO.File.WriteAllText(finalExcelOutputFilePath, worksheetXML);
            }
            catch (Exception ex)
            {
                WriteLine($"{ex}");
                var patVersion = new Regex(@"");
            }
            //WriteLine($"files processed={string.Join(",\n", mapFileNameToColumnPosition.Keys.Select(f=>$"@\"{f}\"").ToList())}");
            var excel = ExpandEnvironmentVariables(@"%MSOFFICE%\EXCEL.EXE");
            System.Diagnostics.Process.Start(excel, $"/s \"{finalExcelOutputFilePath}\"");
        }
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

    static void UpdateMapFileNameToSymbolToDateToAttributes(AutoMultiDimSortedDictionary<string/*file name*/, AutoMultiDimSortedDictionary<string /*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*metric name*/, string/*metric value*/>>>> mapFileNameToSymbolToDatesToAttributes,
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
    static Regex patDateTime_ddd_MMM_dd_YYYY_ddd_MMM_dd_YYYY = new Regex(@"^(Mon|Tue|Wed|Thu|Fri|Sat|Sun) ((Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) ([0-9]+),? ([0-9]+)):\s*&lt;((Mon|Tue|Wed|Thu|Fri|Sat|Sun) ((Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) ([0-9]+),? ([0-9]+)))&gt;");
    static Regex patDateTime_ddd_MMM_dd_YYYY = new Regex(@"^(Mon|Tue|Wed|Thu|Fri|Sat|Sun) ((Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) ([0-9]+),? ([0-9]+))");

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
    static Dictionary<int, (string name,bool numeric)> mapPositionToAttributeName = new Dictionary<int, (string name, bool numeric)>{
        //"Name", 
        { 1, ("Current Price", true)},
        //"Price $ Chg", 
        { 2 , ("Price % Chg",true) },
        //"% Off High",  
        //"Vol % Chg vs 50-Day", "50-Day Avg Vol (1000s)",  "50-Day Avg $ Vol (1000s)",  
        //"Market Cap (mil)",
        { 3, ("Comp Rating", true) },
        { 4, ("EPS Rating", true) },
        { 5 , ("RS Rating", true) },
        { 6 , ("A/D Rating", false) },
        { 7 , ("SMR Rating", false) },
        { 8 , ("50-Day Avg Vol (1000s)", true) },
        { 9, ("Ind Group Rank", true) },
        {10, ("Industry Name", false) },
        {11, ("Sector", false) }
    };
    static Dictionary<string, int> mapAttributeNameToPosition = new();
    static Regex patRestoreCommas = new Regex(@"\&comma;");
    static Regex patRestoreNewLines = new Regex(@"\&(newline|nl);");
    static (Regex regex, string replacement)[] patHighLightKeyWord = {
        (new Regex("NEUTRAL"), """<B><Font html:Size="14" html:Color="#FFC000">NEUTRAL</Font></B>"""),
        (new Regex("(?<!STRONG\\s*)SELL"),    """<B><Font html:Size="14" html:Color="#FF0000">SELL</Font></B>"""),
        (new Regex("Reiterate Buy",RegexOptions.IgnoreCase), """<B><Font html:Color="#008080">Reiterate Buy</Font></B>"""),
        (new Regex("(?<!STRONG\\s*)BUY"),"""<B><Font html:Size="14" html:Color="#008080">BUY</Font></B>"""),
        (new Regex("LOW"),"""<B><Font html:Size="14" html:Color="#008080">LOW</Font></B>"""),
        (new Regex("HIGH"),"""<B><Font html:Size="14" html:Color="#FF0000">HIGH</Font></B>"""),
        (new Regex("STRONG\\s*BUY"),"""<B><I><Font html:Size="15" html:Color="#008080">STRONG BUY</Font></I></B>"""),
        (new Regex("STRONG\\s*SELL"),"""<B><I><Font html:Size="15" html:Color="#FF0000">STRONG SELL</Font></I></B>"""),
        (new Regex("MODERATE"),"""<B><Font html:Size="14" html:Color="#00CCFF">MODERATE</Font></B>"""),
        (new Regex("MEDIUM"),"""<B><Font html:Size="14" html:Color="#00CCFF">MEDIUM</Font></B>"""),
        (new Regex("HOLD"),"""<B><Font html:Size="14" html:Color="#00CCFF">HOLD</Font></B>""")
    };
    static AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*metric name*/, string/*metric value*/>>> ParseCSV(string fileName, DateTime listDateTime)
    {
        var useCfg = true;
        AutoMultiDimSortedDictionary<string/*symbol*/, AutoMultiDimSortedDictionary<DateTime, AutoInitSortedDictionary<string/*metric name*/, string/*metric value*/>>> rows = new();
        var configuration = new CsvHelper.Configuration.CsvConfiguration(new CultureInfo("en-US", false))
        {
            MissingFieldFound = null,
            HasHeaderRecord = true
        };
        //WriteLine($"ParseCSV: open {fileName}");
        //System.Diagnostics.Debugger.Break();
        using (StreamReader reader = new StreamReader(fileName))
        //using (var csv = new CsvHelper.CsvReader(reader,  new CsvHelper.Configuration.CsvConfiguration {  Delimiter = ",",  HasHeaderRecord = true })) {
        //using (CsvReader csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        using (CsvReader csv = useCfg ? new CsvReader(reader, configuration) : new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            //csv.Configuration.HasHeaderRecord = false;

            var isNotesFileName = patNotes.Match(fileName);
            var isSharesFileName = patShares.Match(fileName);
            var isUnrealizedGainName = patUnrealizedGains.Match(fileName);
            //WriteLine($"f={fileName} {isUnrealizedGainName}");
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
                    //if(!isNotesFileName.Success)
                    result = new AutoInitSortedDictionary<string, string>(); // In notes files we can have multiple lines for the same symbol so we use the same results
                    if (isUnrealizedGainName.Success)
                    {
                        name = isUnrealizedGainName.Groups[3].Value;
                        var sharesOrProfitLossOrUnitCost = csv.GetField<string>(name);
                        result.Add(name, sharesOrProfitLossOrUnitCost);
                    }
                    else if (isNotesFileName.Success)
                    {
                        var noteDaysOld = 0.0;
                        name = isNotesFileName.Groups[2].Value;
                        var dateString = csv.GetField<string>("Date");
                        if (DateTime.TryParse(dateString, out DateTime date))
                        {
                            noteDaysOld = (TODAY - date).TotalDays;
                            dateString = date.ToString("ddd MMM dd, yy");
                        }
                        var note = csv.GetField<string>("Notes");
                        note = convertPunctuationsToXML(note, patRestoreCommas, ",");
                        note = convertPunctuationsToXML(note, patRestoreNewLines, "&#10;");
                        if (!fileName.Contains("TRADE"))
                            foreach ((var pat, var replace) in patHighLightKeyWord)
                            {
                                if (pat.Match(note).Success)
                                {
                                    note = pat.Replace(note, replace);
                                    //break;
                                }
                            }
                        notes = dateString + ": " + note;
                        // Separate multiple notes with a new line
                        //var newline = result.ContainsKey($"{name}Notes") ? "&#10;" : "";
                        if (result.ContainsKey($"{name}Notes"))
                        {
                            if (noteDaysOld < maxNotesDaysOld) // @@todo@@ Does this ever execute? I think NOT! I think this is now superfluous since we get a new result every time thru
                            {
                                var oldSize = result[$"{name}Notes"].Length;
                                var newAdditional = note.Length;
                                if (oldSize < maxNoteSize)
                                {
                                    result[$"{name}Notes"] = result[$"{name}Notes"] + "&#10;" + notes;
                                    var newSize = result[$"{name}Notes"].Length;
                                    if(debug1)WriteLine($"""Adding note:      {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld}) oldSize={oldSize} adding={newAdditional} newSize={newSize}""");
                                }
                                else
                                {
                                    if(debug1)WriteLine($"""note too big:     {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld}) oldSize={oldSize} adding={newAdditional}""");
                                    skip = true;
                                }
                            }
                            else
                            {
                                if(debug1)WriteLine($"""Too old:          {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld})""");
                                skip = true;
                            }
                        }
                        else
                        {
                            if (noteDaysOld < maxNotesDaysOld)
                            {
                                if(debug1)WriteLine($"""Initial Add:      {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld})""");
                                result.Add($"{name}Notes", notes);
                            }
                            else
                            {
                                if(debug1) WriteLine($"""Initial Add skip: {name}Notes[{symbol}][{dateString}]=="{note.Substring(0, Math.Min(note.Length, 20))}" days old ({noteDaysOld})""");
                                skip = true;
                            }
                        }
                    }
                    else
                    {
                        //result.Add("Industry Name", csv.GetField<string>("Industry Name"));
                        foreach (var (_, attrName) in mapPositionToAttributeName)
                        {
                            var attrValue = csv.GetField<string>(attrName.Item1);
                            result.Add(attrName.Item1, attrValue);
                        }
                        result.Add("seq", (count - 1).ToString("00")); // this is for lists like IBD 50 where the order is important
                        result.Add("Updated", new FileInfo(fileName).LastWriteTime.ToString("ddd MMM dd yy"));
                    }
                    if (skip)
                    {
                            if(debug1)WriteLine($"adding row SKIP    [{symbol}][{listDateTime}] from file={fileName}");
                    }
                    else if(rows.ContainsKey(symbol) && rows[symbol].ContainsKey(listDateTime))
                    {
                        foreach((var key,var val) in result)
                        {
                            rows[symbol][listDateTime][key] = rows[symbol][listDateTime][key] + "&#10;" + val;
                        }
                    }
                    else
                    {
                        try
                        {
                            rows[symbol][listDateTime] = result;
                            if(debug1) WriteLine($"adding row         [{symbol}][{listDateTime}] from file={fileName}");
                        }
                        catch (System.ArgumentException)
                        {
                            rows[symbol][listDateTime][$"{name}Notes"] = notes + "&#10;" + rows[symbol][listDateTime][$"{name}Notes"];
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

    static List<string> GetFileNamesAndSwitchesFromArgs(string[] args)
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
        return fns;
    }

}
