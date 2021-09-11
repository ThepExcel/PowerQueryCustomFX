(OriginalList as list,PosIndex as list) as list=>
let
    // MCode created by ThepExcel.com
    OriginalListCount=List.Count(OriginalList),
    PosIndexFix=List.Transform(PosIndex,each if _<0 then OriginalListCount+_ else _),
    GetList=List.Transform(PosIndexFix,each OriginalList{_})
in
    GetList
