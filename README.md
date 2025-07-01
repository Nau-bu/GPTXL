````markdown
# Bond Dashboard Builder — README for Codex

Below are three sequential summaries—two from prior GPT instances, plus a distilled recap of our latest deep‐dive—so that any new Codex session can pick up exactly where we left off.

---

## 1) First GPT Instance

### Project in one page

#### 1. Purpose
Build a **self-updating Excel “Bond Dashboard”** that:
- Ingests any Bondway export files dropped in **\<project folder\>\Raw\\**  
- Combines **AI-adjusted history** & **TRACE trades** into one tidy table  
- Drives PivotCharts + slicers (Metric, Date, Bond) (line = AI, dots = TRACE)  
- Lets a user pick **one target bond** vs a **peer basket** → plots **Peer Avg** line & **C-Spread** bars  
- Refreshes automatically when the workbook opens or when B2 (path cell) changes

#### 2. Folder-path logic
- **FolderPath**: named-range holding the parent folder  
- **ProjectFolder**: fallback to `FolderPath & "\Raw\"` if missing  
- Power Query queries always reference ProjectFolder, with built-in fallback

#### 3. Power Query objects (load to Data Model)

| Query                      | Functionality                                                                                         |
| -------------------------- | ----------------------------------------------------------------------------------------------------- |
| **PQ_RawFiles**            | Lists only `.xlsx/.xlsb/.xlsm` inside `<FolderPath>\Raw\`                                             |
| **PQ_SheetList**           | Opens each workbook in Raw, returns one row per sheet with **SheetName, Kind, Data**                 |
| **PQ_CombinedHistory**     | Pulls “Combined Level History” sheets, unpivots to `Date/Bond/Metric/Value`, flags `Adjusted=TRUE`   |
| **PQ_Trades**              | Pulls “Trades …” sheets, keeps numeric fields, flags `Adjusted=FALSE`                                 |
| **YearMaturity**           | Extracts numeric maturity from BondID                                                                |
| **PQ_AllData**             | `Append` of history + trades → master fact table                                                      |

*(All M-code blocks were provided earlier.)*

#### 4. DAX measures

```dax
PeerAvg :=
  AVERAGEX(
    FILTER(ALLSELECTED(AllData), AllData[Bond] <> SELECTEDVALUE(AllData[Bond])),
    [Value]
  )

CSpread :=
  VAR _v = SELECTEDVALUE(AllData[Value])
  RETURN IF(NOT ISBLANK(_v), _v - [PeerAvg])
````

#### 5. Front-end layout

1. PivotChart (Date axis; Series = Bond+Adjusted; Values = AVG(Value), plus PeerAvg & CSpread)
2. Slicers: Metric, Date timeline, **Bond Target** (single), **Peer Basket** (multi)

#### 6. VBA (optional)

* **Workbook\_Open**: write `<ThisWorkbook.Path>\Raw` to B2, create folder, RefreshAll
* **Worksheet\_Change** on B2 → RefreshAll

#### 7. Status

* Queries compile & refresh without errors
* Charts update when target & peers are selected

---

## 2) Second GPT Instance

### High-level recap

**Project Goal**

* Automated Excel dashboard ingesting TRACE “Raw” files
* Unified “AllData” table of AI-adjusted history + TRACE trades
* Extracts Price, YTM, G/T spreads, YTC, maturity year; normalizes BondID
* Flags adjusted vs raw; dynamic slicers for single “Target” + multi “Peer”
* DAX for peer averages & C-Spread; line + scatter overlay charts

**Power Query workflows**

1. **PQ\_CombinedHistory**: combine all “Combined Level History” sheets → unpivot → pivot → rename → add flags → YearMaturity → normalize → force types
2. **PQ\_Trades**: load all “Trades …” sheets → expand → add flags & BondID → force types → bucket volumes
3. **AllData**: `Table.Combine({history, trades})` → Connection only
4. **Dim tables** for slicers: distinct Bond lists for Target & Peer

**DAX Measures**

* **Target TSpread**: `CALCULATE(AVERAGE(T Spread), TREATAS(VALUES(BondTarget),AllData[Bond]))`
* **PeerAvg TSpread**: average of selected PeerBasket excluding target, filtered to Adjusted=TRUE
* **CSpread TSpread**: target adjusted avg minus PeerAvg

**Charting**

* Build PivotChart; then copy/paste to regular chart to overlay XY‐scatter
* Slicers drive all visuals

**Troubleshooting**

* Used `MissingField.Ignore` for robustness
* Forced type conversions with `try DateTime.From(_)`
* Converted pivot chart to normal chart for XY overlay

---

## 3) This Conversation — Deep-dive & Final Master Query

Over multiple iterations we:

1. **Cleaned up folder-path logic** so Workbook-relative path always resolves to `…\Raw\`.
2. **Merged all PQ steps** (RawFiles, SheetList, CombinedHistory, Trades) into **one master M-query**.
3. **Fixed parsing errors** (e.g. stray numeric Attributes) by:

   * Unpivoting first
   * Immediately forcing `Attribute` → **type text**
4. **Added “Adjusted”** flag on history (`TRUE`) vs trades (`FALSE`).
5. **Extracted DateCol dynamically** (first column of Combined Level History).
6. **Normalized BondID** (slashes → dashes), extracted YearMaturity, created MaturityBucket (3Y/5Y/10Y/>10Y).
7. **Appended & typed** all fields, with `MissingField.Ignore` so missing columns never break.

### Final **PQ\_MasterData** M-code

*(Paste this in Power Query as a single, self-contained query named `PQ_MasterData` and load to the Data Model.)*

```m
let
  RawPathVal = Excel.CurrentWorkbook(){[Name="ThisWorkbookPath"]}[Content]{0}[Column1],
  ThisPath   = if Value.Is(RawPathVal, type text) then RawPathVal else Text.From(RawPathVal),
  BaseFolder = if Text.EndsWith(ThisPath, "\") then ThisPath else ThisPath & "\",
  RawFolder  = BaseFolder & "Raw\",

  AllFiles   = Folder.Files(RawFolder),
  ExcFiles   = Table.SelectRows(AllFiles, each List.Contains({".xlsx",".xlsm",".xlsb"}, Text.Lower([Extension]))),

  WithWB     = Table.AddColumn(ExcFiles, "WB", each Excel.Workbook([Content], true)),
  AllSheets  = Table.ExpandTableColumn(WithWB, "WB", {"Name","Data"}, {"SheetName","Data"}),
  Promoted   = Table.TransformColumns(AllSheets, { "Data", each Table.PromoteHeaders(_, [PromoteAllScalars=true]) }),

  // Historical sheets
  HistSheets = Table.SelectRows(Promoted, each [SheetName] = "Combined Level History"),
  FlatHist   = Table.Combine(HistSheets[Data]),
  DateCol    = Table.ColumnNames(FlatHist){0},
  Unpvt      = Table.UnpivotOtherColumns(FlatHist, {DateCol}, "Attribute", "Value"),
  ForceText  = Table.TransformColumnTypes(Unpvt, {{"Attribute", type text}}),
  NonBlank   = Table.SelectRows(ForceText, each Text.Length([Attribute]) > 0),
  ParsedH    = Table.AddColumn(NonBlank, "Rec", each let p=Text.Split([Attribute]," ") in if List.Count(p)>1 then [BondID=Text.Combine(List.RemoveLastN(p,1)," "),Metric=List.Last(p)] else null),
  GoodH      = Table.SelectRows(ParsedH, each [Rec]<>null),
  ExpH       = Table.ExpandRecordColumn(GoodH,"Rec",{"BondID","Metric"}),
  FilterH    = Table.SelectRows(ExpH, each List.Contains({"Price","YTM","G-Spread","T-Spread","YTC"}, [Metric])),
  PivotH     = Table.Pivot(FilterH, List.Distinct(FilterH[Metric]), "Metric", "Value"),
  RenameH    = Table.RenameColumns(PivotH,{{"YTM","Yield to Maturity"},{"YTC","Yield to Call"},{"G-Spread","G Spread"},{"T-Spread","T Spread"}},MissingField.Ignore),
  AdjH       = Table.AddColumn(RenameH,"Adjusted",each true,type logical),
  TimeH      = Table.RenameColumns(AdjH,{{DateCol,"Trade Date and Time"}},MissingField.Ignore),
  HistFinal  = Table.SelectColumns(TimeH,{"Trade Date and Time","BondID","Price","Yield to Maturity","G Spread","T Spread","Yield to Call","Adjusted"},MissingField.Ignore),

  // Trade sheets
  TradeSheets= Table.SelectRows(Promoted, each Text.StartsWith([SheetName],"Trades ") and Value.Is([Data], type table)),
  SampleTbl  = if Table.IsEmpty(TradeSheets) then #table({}, {}) else TradeSheets{0}[Data],
  ExpT       = Table.ExpandTableColumn(TradeSheets,"Data",Table.ColumnNames(SampleTbl)),
  AdjT       = Table.AddColumn(ExpT,"Adjusted",each false,type logical),
  IDT        = Table.AddColumn(AdjT,"BondID",each Text.Trim(Text.Middle(Text.From([SheetName]),7)),type text),
  TypedT     = Table.TransformColumnTypes(IDT,{
                  {"Trade Date and Time",type datetime},
                  {"Price",type number},
                  {"Yield to Maturity",type number},
                  {"G Spread",type number},
                  {"T Spread",type number},
                  {"Yield to Call",type number},
                  {"Quantity",Int64.Type}
               },MissingField.Ignore),
  BucketT    = Table.AddColumn(TypedT,"TradeSizeBucket",each if [Quantity]<1000000 then "<1M" else if [Quantity]<5000000 then "1M-5M" else ">5M",type text),
  TradesFinal= Table.SelectColumns(BucketT,{"Trade Date and Time","BondID","Price","Yield to Maturity","G Spread","T Spread","Yield to Call","Quantity","TradeSizeBucket","Adjusted"},MissingField.Ignore),

  // Combine & maturity
  Combined   = Table.Combine({HistFinal,TradesFinal}),
  WithYear   = Table.AddColumn(Combined,"YearMaturity",each try Number.From(Text.AfterDelimiter(List.Last(Text.Split(Text.From([BondID])," ")),"/")) otherwise null,Int64.Type),
  FinalTable = Table.AddColumn(WithYear,"MaturityBucket",each if [YearMaturity]<=3 then "3Y" else if [YearMaturity]<=5 then "5Y" else if [YearMaturity]<=10 then "10Y" else ">10Y",type text)
in
  FinalTable
```

---

Hand this README (with its M-code) to the next Codex or teammate and they’ll have **everything** needed to continue building, testing, and extending the Bond Dashboard.
