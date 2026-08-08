(TableName as table,ColumnName as any,optional MoveFirst as logical) as table=>
let
    // MCode created by ThepExcel.com
    MoveFirstFix=if MoveFirst is null then true else MoveFirst,
    ColumnNameList=if Value.Type(ColumnName)= type text then {ColumnName} else ColumnName,
    OriginalAllColumnName=Table.ColumnNames(TableName),
    RemoveColResult=List.Difference(OriginalAllColumnName,ColumnNameList),
    Result = 
    if MoveFirstFix=false then
    Table.SelectColumns(TableName,RemoveColResult&ColumnNameList)
    else
    Table.SelectColumns(TableName,ColumnNameList&RemoveColResult)

in
    Result