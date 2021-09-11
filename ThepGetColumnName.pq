(TableName as table,ColNumber as number) as text=>
// MCode created by ThepExcel.com
let
    ColList = Table.ColumnNames(TableName),
    GetOldName=ColList{ColNumber-1}   
in
    GetOldName
