(TableName as table,StartColumnNum as number,optional EndColumnNum as number) as table=>
let
    // MCode created by ThepExcel.com
    EndColumnNum=if EndColumnNum is null then StartColumnNum else EndColumnNum,
    NumColumn=EndColumnNum-StartColumnNum+1,
    Result = Table.SelectColumns(TableName,List.Range(Table.ColumnNames(TableName),StartColumnNum,NumColumn) )

in
    Result