(TableName as table,FindColText as text,optional MatchType as text,optional ColtoKeep as list) as table=>
let
    // MCode created by ThepExcel.com
    MyHeader =   
    if MatchType ="start" then
    List.Select(Table.ColumnNames(TableName),each Text.StartsWith(_,FindColText,Comparer.OrdinalIgnoreCase))  else

    if MatchType ="end" then
    List.Select(Table.ColumnNames(TableName),each Text.EndsWith(_,FindColText,Comparer.OrdinalIgnoreCase)) else
    
    List.Select(Table.ColumnNames(TableName),each Text.Contains(_,FindColText,Comparer.OrdinalIgnoreCase))
    ,

    Result = 
    if ColtoKeep is null then 
    Table.SelectColumns(TableName, MyHeader )
    else
    Table.SelectColumns(TableName, ColtoKeep & MyHeader )
in
    Result