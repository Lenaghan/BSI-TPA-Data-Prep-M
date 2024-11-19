let
  Source = All, 
  #"Filtered Rows" = Table.SelectRows(Source, each ([LOB] = "VA")), 
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
    {"Driver", "Date of Loss", "Accident State"}, 
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
      {"Loc Description", each List.Max([Loc Description]), type nullable text}, 
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
  #"Replace VA with AL" = Table.ReplaceValue(
    #"Capitalized Each Word", 
    "VA", 
    "AL", 
    Replacer.ReplaceText, 
    {"LOB"}
  ), 
  #"Merged Queries" = Table.NestedJoin(
    #"Replace VA with AL", 
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
      "Non-Legal Recovery"
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
      "Claim Description", 
      "Driver"
    }
  ), 
  #"Sorted Rows" = Table.Sort(#"Reordered Columns", {{"Total Incurred", Order.Descending}})
in
  #"Sorted Rows"
