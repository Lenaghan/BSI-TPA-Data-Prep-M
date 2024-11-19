// All
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

// AsOf
let
  FilePath = Excel.CurrentWorkbook(){[Name = "FilePathFolder"]}[Content]{0}[Column1], 
  FileName = Excel.CurrentWorkbook(){[Name = "FileName"]}[Content]{0}[Column1], 
  Source = Excel.Workbook(File.Contents(FilePath & FileName), null), 
  #"Prompt Summary_Sheet" = Source{[Item = "Prompt Summary", Kind = "Sheet"]}[Data], 
  #"Removed Columns" = Table.RemoveColumns(#"Prompt Summary_Sheet", {"Column2", "Column3"}), 
  #"Split Column by Delimiter" = Table.SplitColumn(
    #"Removed Columns", 
    "Column1", 
    Splitter.SplitTextByEachDelimiter({":"}, QuoteStyle.Csv, false), 
    {"Column1.1", "Column1.2"}
  ), 
  #"Removed Top Rows" = Table.Skip(#"Split Column by Delimiter", 1), 
  #"Kept First Rows" = Table.FirstN(#"Removed Top Rows", 1), 
  #"Transposed Table" = Table.Transpose(#"Kept First Rows"), 
  #"Promoted Headers" = Table.PromoteHeaders(#"Transposed Table", [PromoteAllScalars = true]), 
  #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {{"Valuation Date", type date}})
in
  #"Changed Type"

// CGL
let
  Source = All, 
  #"Filtered Rows" = Table.SelectRows(Source, each ([LOB] = "GC")), 
  #"Changed Type" = Table.TransformColumnTypes(
    #"Filtered Rows", 
    {
      {"Total Incurred", Currency.Type}, 
      {"Total Recovery", Currency.Type}, 
      {"Total O/S", Currency.Type}, 
      {"Total Paid", Currency.Type}, 
      {"Non-Legal Inc", Currency.Type}, 
      {"Legal Inc", Currency.Type}, 
      {"Non-Legal Recovery", Currency.Type}, 
      {"Legal Recovery", Currency.Type}, 
      {"Non-Legal Outsd", Currency.Type}, 
      {"Legal Outsd", Currency.Type}, 
      {"Non-Legal PTD", Currency.Type}, 
      {"Legal PTD", Currency.Type}, 
      {"Bod Inj Incurred", Currency.Type}, 
      {"Bod Inj Recovery", Currency.Type}, 
      {"Bod Inj O/S", Currency.Type}, 
      {"Bod Inj Paid", Currency.Type}, 
      {"Prop Dmg Incurred", Currency.Type}, 
      {"Prop Dmg Recovery", Currency.Type}, 
      {"Prop Dmg O/S", Currency.Type}, 
      {"Prop Dmg Paid", Currency.Type}, 
      {"Effective", type date}, 
      {"Expiration", type date}, 
      {"Date of Loss", type date}
    }
  ), 
  #"Group By IVD, DOL, State" = Table.Group(
    #"Changed Type", 
    {"Loc Description", "Date of Loss", "Accident State"}, 
    {
      {"Effective", each List.Max([Effective]), type nullable date}, 
      {"Expiration", each List.Max([Expiration]), type nullable date}, 
      {"Occurrence Number", each Text.Combine(List.Distinct([Occurrence Number]), "; ")}, 
      {"Carrier", each List.Max([Carrier]), type nullable text}, 
      {"LOB", each List.Max([LOB]), type nullable text}, 
      {"Claim Number", each Text.Combine(List.Distinct([Claim Number]), "; ")}, 
      {"Claimant", each Text.Combine(List.Distinct([Claimant]), "; ")}, 
      {"Claim Status", each List.Max([Claim Status]), type nullable text}, 
      {"Prop Dmg Paid", each List.Sum([Prop Dmg Paid]), type number}, 
      {"Prop Dmg O/S", each List.Sum([#"Prop Dmg O/S"]), type number}, 
      {"Prop Dmg Recovery", each List.Sum([Prop Dmg Recovery]), type number}, 
      {"Bod Inj Paid", each List.Sum([Bod Inj Paid]), type number}, 
      {"Bod Inj O/S", each List.Sum([#"Bod Inj O/S"]), type number}, 
      {"Bod Inj Recovery", each List.Sum([Bod Inj Recovery]), type number}, 
      {"Legal PTD", each List.Sum([Legal PTD]), type number}, 
      {"Non-Legal PTD", each List.Sum([#"Non-Legal PTD"]), type number}, 
      {"Legal Outsd", each List.Sum([Legal Outsd]), type number}, 
      {"Non-Legal Outsd", each List.Sum([#"Non-Legal Outsd"]), type number}, 
      {"Legal Recovery", each List.Sum([Legal Recovery]), type number}, 
      {"Non-Legal Recovery", each List.Sum([#"Non-Legal Recovery"]), type number}, 
      {"Division", each List.Max([Division]), type nullable text}, 
      {"Loc Code", each List.Max([Loc Code]), type nullable text}, 
      {"Driver", each List.Max([Driver]), type nullable text}, 
      {"Claim Description", each List.Max([Claim Description]), type nullable text}
    }
  ), 
  #"Added Custom" = Table.AddColumn(
    #"Group By IVD, DOL, State", 
    "Policy Year", 
    each Text.From([Effective]) & "-" & Text.End(Text.From(Date.Year([Expiration])), 2)
  ), 
  #"Add ALAE Paid" = Table.AddColumn(
    #"Added Custom", 
    "ALAE Paid", 
    each [Legal PTD] + [#"Non-Legal PTD"]
  ), 
  #"Add ALAE O/S" = Table.AddColumn(
    #"Add ALAE Paid", 
    "ALAE O/S", 
    each [Legal Outsd] + [#"Non-Legal Outsd"]
  ), 
  #"Add ALAE Recovery" = Table.AddColumn(
    #"Add ALAE O/S", 
    "ALAE Recovery", 
    each [Legal Recovery] + [#"Non-Legal Recovery"]
  ), 
  #"Add Prop Dmg Incurred" = Table.AddColumn(
    #"Add ALAE Recovery", 
    "Prop Dmg Incurred", 
    each [Prop Dmg Paid] + [#"Prop Dmg O/S"] + [Prop Dmg Recovery]
  ), 
  #"Add Bod Inj Incurred" = Table.AddColumn(
    #"Add Prop Dmg Incurred", 
    "Bod Inj Incurred", 
    each [Bod Inj Paid] + [#"Bod Inj O/S"] + [Bod Inj Recovery]
  ), 
  #"Add ALAE Incurred" = Table.AddColumn(
    #"Add Bod Inj Incurred", 
    "ALAE Incurred", 
    each [ALAE Paid] + [#"ALAE O/S"] + [ALAE Recovery]
  ), 
  #"Add Total Paid" = Table.AddColumn(
    #"Add ALAE Incurred", 
    "Total Paid", 
    each [Prop Dmg Paid] + [Bod Inj Paid] + [ALAE Paid]
  ), 
  #"Add Total O/S" = Table.AddColumn(
    #"Add Total Paid", 
    "Total O/S", 
    each [#"Prop Dmg O/S"] + [#"Bod Inj O/S"] + [#"ALAE O/S"]
  ), 
  #"Add Total Recovery" = Table.AddColumn(
    #"Add Total O/S", 
    "Total Recovery", 
    each [Prop Dmg Recovery] + [Bod Inj Recovery] + [ALAE Recovery]
  ), 
  #"Add Net Paid" = Table.AddColumn(
    #"Add Total Recovery", 
    "Net Paid", 
    each [Total Paid] + [Total Recovery]
  ), 
  #"Add Total Incurred" = Table.AddColumn(
    #"Add Net Paid", 
    "Total Incurred", 
    each [Net Paid] + [#"Total O/S"]
  ), 
  #"Add Net Paid (excl. ALAE)" = Table.AddColumn(
    #"Add Total Incurred", 
    "Net Paid (excl. ALAE)", 
    each [Prop Dmg Paid] + [Prop Dmg Recovery] + [Bod Inj Paid] + [Bod Inj Recovery]
  ), 
  #"Add Tot Incurred (excl. ALAE)" = Table.AddColumn(
    #"Add Net Paid (excl. ALAE)", 
    "Tot Incurred (excl. ALAE)", 
    each [#"Net Paid (excl. ALAE)"] + [#"Prop Dmg O/S"] + [#"Bod Inj O/S"]
  ), 
  #"Capitalized Each Word" = Table.TransformColumns(
    #"Add Tot Incurred (excl. ALAE)", 
    {
      {"Driver", Text.Proper, type text}, 
      {"Carrier", Text.Proper, type text}, 
      {"Claimant", Text.Proper, type text}, 
      {"Division", Text.Proper, type text}, 
      {"Loc Description", Text.Proper, type text}
    }
  ), 
  #"Replace GC with CGL" = Table.ReplaceValue(
    #"Capitalized Each Word", 
    "GC", 
    "CGL", 
    Replacer.ReplaceText, 
    {"LOB"}
  ), 
  #"Merged Queries" = Table.NestedJoin(
    #"Replace GC with CGL", 
    {"Effective", "LOB"}, 
    Deductibles, 
    {"Effective", "LOB"}, 
    "Deductibles", 
    JoinKind.LeftOuter
  ), 
  #"Expanded Deductibles" = Table.ExpandTableColumn(
    #"Merged Queries", 
    "Deductibles", 
    {"Deductible"}, 
    {"Deductible"}
  ), 
  #"Add Limited Paid" = Table.AddColumn(
    #"Expanded Deductibles", 
    "Limited Paid", 
    each List.Min({[Net Paid], [Deductible]})
  ), 
  #"Add Limited Incurred" = Table.AddColumn(
    #"Add Limited Paid", 
    "Limited Incurred", 
    each List.Min({[Total Incurred], [Deductible]})
  ), 
  #"Add Limited O/S" = Table.AddColumn(
    #"Add Limited Incurred", 
    "Limited O/S", 
    each [Limited Incurred] - [Limited Paid]
  ), 
  #"Add Excess Paid" = Table.AddColumn(
    #"Add Limited O/S", 
    "Excess Paid", 
    each [Net Paid] - [Limited Paid]
  ), 
  #"Add Excess Incurred" = Table.AddColumn(
    #"Add Excess Paid", 
    "Excess Incurred", 
    each [Total Incurred] - [Limited Incurred]
  ), 
  #"Add Excess O/S" = Table.AddColumn(
    #"Add Excess Incurred", 
    "Excess O/S", 
    each [Excess Incurred] - [Excess Paid]
  ), 
  #"Removed Columns" = Table.RemoveColumns(
    #"Add Excess O/S", 
    {
      "Effective", 
      "Expiration", 
      "Carrier", 
      "Legal PTD", 
      "Non-Legal PTD", 
      "Legal Outsd", 
      "Non-Legal Outsd", 
      "Legal Recovery", 
      "Non-Legal Recovery", 
      "Driver"
    }
  ), 
  #"Reordered Columns" = Table.ReorderColumns(
    #"Removed Columns", 
    {
      "LOB", 
      "Policy Year", 
      "Occurrence Number", 
      "Claim Number", 
      "Date of Loss", 
      "Claimant", 
      "Accident State", 
      "Claim Status", 
      "Prop Dmg Paid", 
      "Prop Dmg O/S", 
      "Prop Dmg Recovery", 
      "Prop Dmg Incurred", 
      "Bod Inj Paid", 
      "Bod Inj O/S", 
      "Bod Inj Recovery", 
      "Bod Inj Incurred", 
      "ALAE Paid", 
      "ALAE O/S", 
      "ALAE Recovery", 
      "ALAE Incurred", 
      "Total Paid", 
      "Total O/S", 
      "Total Recovery", 
      "Net Paid", 
      "Total Incurred", 
      "Net Paid (excl. ALAE)", 
      "Tot Incurred (excl. ALAE)", 
      "Deductible", 
      "Limited Paid", 
      "Limited O/S", 
      "Limited Incurred", 
      "Excess Paid", 
      "Excess O/S", 
      "Excess Incurred", 
      "Division", 
      "Loc Code", 
      "Loc Description", 
      "Claim Description"
    }
  ), 
  #"Sorted Rows" = Table.Sort(#"Reordered Columns", {{"Total Incurred", Order.Descending}})
in
  #"Sorted Rows"

// Deductibles
let
  Source = Excel.CurrentWorkbook(){[Name = "Deductibles"]}[Content], 
  #"Changed Type" = Table.TransformColumnTypes(
    Source, 
    {{"Effective", type date}, {"Deductible", Currency.Type}, {"LOB", type text}}
  )
in
  #"Changed Type"
