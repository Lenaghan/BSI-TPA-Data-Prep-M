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
