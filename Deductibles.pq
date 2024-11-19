let
  Source = Excel.CurrentWorkbook(){[Name = "Deductibles"]}[Content], 
  #"Changed Type" = Table.TransformColumnTypes(
    Source, 
    {{"Effective", type date}, {"Deductible", Currency.Type}, {"LOB", type text}}
  )
in
  #"Changed Type"
