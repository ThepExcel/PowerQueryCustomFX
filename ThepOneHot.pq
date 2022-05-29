(TableName as table,TargetColumnName as text) as table=>
let

    Source = TableName,
    #"Added Index" = Table.AddIndexColumn(Source, "ThepIndex", 0, 1, Int64.Type),
    #"Duplicated Column" = Table.DuplicateColumn(#"Added Index", TargetColumnName, TargetColumnName&" - Copy"),
    TargetList=Table.ToList(Table.SelectColumns(Source,TargetColumnName)),
    #"Pivoted Column" = Table.Pivot(#"Duplicated Column", List.Distinct(TargetList), TargetColumnName&" - Copy", "ThepIndex", List.NonNullCount)

in
    #"Pivoted Column"
