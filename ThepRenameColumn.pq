(TableName as table,ColNumber as any,NewName as any) as table=>

let
    // MCode created by ThepExcel.com

    GetMultipleListItem=(OriginalList as list,PosIndex as list)=>
    let
        OriginalListCount=List.Count(OriginalList),
        PosIndexFix=List.Transform(PosIndex,each if _<0 then OriginalListCount+_ else _-1),
        GetList=List.Transform(PosIndexFix,each OriginalList{_})
    in
        GetList
    ,
    //
    ColNumberList=if Value.Type(ColNumber)=type number then {ColNumber} else ColNumber,
    NewNameList=if Value.Type(NewName)=type text then {NewName} else NewName,
    ColList = Table.ColumnNames(TableName),    
    GetOldName=GetMultipleListItem(ColList,ColNumberList),
    FinalNameZip=List.Zip({GetOldName,NewNameList}),
    NewNameFix=Table.RenameColumns(TableName,FinalNameZip)       
in
    NewNameFix
