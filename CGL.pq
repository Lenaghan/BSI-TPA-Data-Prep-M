
let
  FilePath = Excel.CurrentWorkbook(){[Name = "FilePathFolder"]}[Content]{0}[Column1], 
  FileName = Excel.CurrentWorkbook(){[Name = "FileName"]}[Content]{0}[Column1], 
  Source = Excel.Workbook(File.Contents(FilePath & FileName), null), 
  RTAD310_Sheet = Source{[Item = "RTAD310", Kind = "Sheet"]}[Data], 
  #"Clean & Promote Headers" = 
    let
      firstN = Table.FirstN(RTAD310_Sheet, 1), 
      renames = List.Transform(
        Table.ColumnNames(RTAD310_Sheet), 
        each {
          _, 
          Text.Combine(
            List.Transform(
              Table.Column(firstN, _), 
              each Text.Trim(Text.Replace(Text.From(_), "#(lf)", " "))
            ), 
            ""
          )
        }
      ), 
      renamedTable = Table.RenameColumns(RTAD310_Sheet, renames)
    in
      Table.Skip(renamedTable, 1), 
  #"Split Column by Delimiter" = Table.SplitColumn(
    #"Clean & Promote Headers", 
    "Program Name", 
    Splitter.SplitTextByEachDelimiter({" - "}, QuoteStyle.Csv, false), 
    {"Client", "Program Name.2"}
  ), 
  #"Separate Eff & Exp dates" = Table.SplitColumn(
    #"Split Column by Delimiter", 
    "Experience Period", 
    Splitter.SplitTextByEachDelimiter({"-"}, QuoteStyle.Csv, false), 
    {"Effective", "Expiration"}
  ), 
  #"Set Date Types" = Table.TransformColumnTypes(
    #"Separate Eff & Exp dates", 
    {{"Loss Date", type date}, {"Effective", type date}, {"Expiration", type date}}
  ), 
  #"Full Outer Join with AsOf" = Table.NestedJoin(
    #"Set Date Types", 
    {"Loss Date"}, 
    AsOf, 
    {"Valuation Date"}, 
    "AsOf", 
    JoinKind.FullOuter
  ), 
  #"Expanded AsOf" = Table.ExpandTableColumn(
    #"Full Outer Join with AsOf", 
    "AsOf", 
    {"Valuation Date"}, 
    {"Valuation Date"}
  ), 
  #"Filled Up" = Table.FillUp(#"Expanded AsOf", {"Valuation Date"}), 
  #"Removed Coll & Incident Only" = Table.SelectRows(
    #"Filled Up", 
    each (
      [Minor Coverage Description]
        <> "Commercial Auto - Collision"
        and [Minor Coverage Description] <> "Commercial Auto - Incident Report"
    )
  ), 
  #"Grouped by Occurrence" = Table.Group(
    #"Removed Coll & Incident Only", 
    {"Occurrence Number"}, 
    {
      {"Carrier", each List.Max([Insurer Name]), type nullable text}, 
      {"LOB", each List.Max([LOB Code]), type nullable text}, 
      {"Effective", each List.Max([Effective]), type nullable date}, 
      {"Expiration", each List.Max([Expiration]), type nullable date}, 
      {"Claim Number", each Text.Combine(List.Distinct([Claim Number]), "; ")}, 
      {"Date of Loss", each List.Max([Loss Date]), type nullable date}, 
      {"Claimant", each Text.Combine(List.Distinct([Claimant Name]), "; ")}, 
      {"Accident State", each List.Max([Loss State]), type nullable text}, 
      {"Claim Status", each List.Max([Claim Status Description]), type nullable text}, 
      {"Prop Dmg Paid", each List.Sum([#"Ind/PD PTD"]), type number}, 
      {"Prop Dmg O/S", each List.Sum([#"Ind/PD Outsd"]), type number}, 
      {"Prop Dmg Recovery", each List.Sum([#"Ind/PD Recov"]), type number}, 
      {"Prop Dmg Incurred", each List.Sum([#"Ind/PD Inc"]), type number}, 
      {"Bod Inj Paid", each List.Sum([#"Med/BI PTD"]), type number}, 
      {"Bod Inj O/S", each List.Sum([#"Med/BI Outsd"]), type number}, 
      {"Bod Inj Recovery", each List.Sum([#"Med/BI Recov"]), type number}, 
      {"Bod Inj Incurred", each List.Sum([#"Med/BI Inc"]), type number}, 
      {"Legal PTD", each List.Sum([Legal PTD]), type number}, 
      {"Non-Legal PTD", each List.Sum([#"Non- Legal PTD"]), type number}, 
      {"Legal Outsd", each List.Sum([Legal Outsd]), type number}, 
      {"Non-Legal Outsd", each List.Sum([#"Non- Legal Outsd"]), type number}, 
      {"Legal Recovery", each List.Sum([Legal Recov]), type number}, 
      {"Non-Legal Recovery", each List.Sum([#"Non-Legal Recov"]), type number}, 
      {"Legal Inc", each List.Sum([Legal Inc]), type number}, 
      {"Non-Legal Inc", each List.Sum([#"Non- Legal Inc"]), type number}, 
      {"Total Paid", each List.Sum([Total PTD]), type number}, 
      {"Total O/S", each List.Sum([Total Outsd]), type number}, 
      {"Total Recovery", each List.Sum([Total Recov]), type number}, 
      {"Total Incurred", each List.Sum([#"Total Inc (Gross)"]), type number}, 
      {"Division", each List.Max([Level 2 Name]), type nullable text}, 
      {"Loc Code", each List.Max([Location Code]), type nullable text}, 
      {"Loc Description", each List.Max([Location Description]), type nullable text}, 
      {"Claim Description", each List.Max([Cause Description]), type nullable text}, 
      {"Driver", each List.Max([Insured  Driver Last Name]), type nullable text}
    }
  ), 
  #"Trimmed Text" = Table.TransformColumns(
    Table.TransformColumnTypes(
      #"Grouped by Occurrence", 
      {
        {"Prop Dmg Paid", type text}, 
        {"Prop Dmg O/S", type text}, 
        {"Prop Dmg Recovery", type text}, 
        {"Prop Dmg Incurred", type text}, 
        {"Bod Inj Paid", type text}, 
        {"Bod Inj O/S", type text}, 
        {"Bod Inj Recovery", type text}, 
        {"Bod Inj Incurred", type text}, 
        {"Legal PTD", type text}, 
        {"Non-Legal PTD", type text}, 
        {"Legal Outsd", type text}, 
        {"Non-Legal Outsd", type text}, 
        {"Legal Recovery", type text}, 
        {"Non-Legal Recovery", type text}, 
        {"Legal Inc", type text}, 
        {"Non-Legal Inc", type text}, 
        {"Total Paid", type text}, 
        {"Total O/S", type text}, 
        {"Total Recovery", type text}, 
        {"Total Incurred", type text}
      }, 
      "en-US"
    ), 
    {
      {"Occurrence Number", Text.Trim, type text}, 
      {"Carrier", Text.Trim, type text}, 
      {"LOB", Text.Trim, type text}, 
      {"Claim Number", Text.Trim, type text}, 
      {"Claimant", Text.Trim, type text}, 
      {"Accident State", Text.Trim, type text}, 
      {"Claim Status", Text.Trim, type text}, 
      {"Prop Dmg Paid", Text.Trim, type text}, 
      {"Prop Dmg O/S", Text.Trim, type text}, 
      {"Prop Dmg Recovery", Text.Trim, type text}, 
      {"Prop Dmg Incurred", Text.Trim, type text}, 
      {"Bod Inj Paid", Text.Trim, type text}, 
      {"Bod Inj O/S", Text.Trim, type text}, 
      {"Bod Inj Recovery", Text.Trim, type text}, 
      {"Bod Inj Incurred", Text.Trim, type text}, 
      {"Legal PTD", Text.Trim, type text}, 
      {"Non-Legal PTD", Text.Trim, type text}, 
      {"Legal Outsd", Text.Trim, type text}, 
      {"Non-Legal Outsd", Text.Trim, type text}, 
      {"Legal Recovery", Text.Trim, type text}, 
      {"Non-Legal Recovery", Text.Trim, type text}, 
      {"Legal Inc", Text.Trim, type text}, 
      {"Non-Legal Inc", Text.Trim, type text}, 
      {"Total Paid", Text.Trim, type text}, 
      {"Total O/S", Text.Trim, type text}, 
      {"Total Recovery", Text.Trim, type text}, 
      {"Total Incurred", Text.Trim, type text}, 
      {"Division", Text.Trim, type text}, 
      {"Loc Code", Text.Trim, type text}, 
      {"Loc Description", Text.Trim, type text}, 
      {"Claim Description", Text.Trim, type text}, 
      {"Driver", Text.Trim, type text}
    }
  ), 
  #"Cleaned Text" = Table.TransformColumns(
    #"Trimmed Text", 
    {
      {"Occurrence Number", Text.Clean, type text}, 
      {"Carrier", Text.Clean, type text}, 
      {"LOB", Text.Clean, type text}, 
      {"Claim Number", Text.Clean, type text}, 
      {"Claimant", Text.Clean, type text}, 
      {"Accident State", Text.Clean, type text}, 
      {"Claim Status", Text.Clean, type text}, 
      {"Prop Dmg Paid", Text.Clean, type text}, 
      {"Prop Dmg O/S", Text.Clean, type text}, 
      {"Prop Dmg Recovery", Text.Clean, type text}, 
      {"Prop Dmg Incurred", Text.Clean, type text}, 
      {"Bod Inj Paid", Text.Clean, type text}, 
      {"Bod Inj O/S", Text.Clean, type text}, 
      {"Bod Inj Recovery", Text.Clean, type text}, 
      {"Bod Inj Incurred", Text.Clean, type text}, 
      {"Legal PTD", Text.Clean, type text}, 
      {"Non-Legal PTD", Text.Clean, type text}, 
      {"Legal Outsd", Text.Clean, type text}, 
      {"Non-Legal Outsd", Text.Clean, type text}, 
      {"Legal Recovery", Text.Clean, type text}, 
      {"Non-Legal Recovery", Text.Clean, type text}, 
      {"Legal Inc", Text.Clean, type text}, 
      {"Non-Legal Inc", Text.Clean, type text}, 
      {"Total Paid", Text.Clean, type text}, 
      {"Total O/S", Text.Clean, type text}, 
      {"Total Recovery", Text.Clean, type text}, 
      {"Total Incurred", Text.Clean, type text}, 
      {"Division", Text.Clean, type text}, 
      {"Loc Code", Text.Clean, type text}, 
      {"Loc Description", Text.Clean, type text}, 
      {"Claim Description", Text.Clean, type text}, 
      {"Driver", Text.Clean, type text}
    }
  ), 
  #"Remove rows before AXA XL" = Table.SelectRows(
    #"Cleaned Text", 
    each [Effective] >= #date(2007, 7, 1)
  )
in
  #"Remove rows before AXA XL"
