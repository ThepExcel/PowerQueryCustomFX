(StartDate as date,EndDate as date) as table=>
let
    // MCode created by ThepExcel.com
    StartDateNum=Number.From(StartDate),
    EndDateNum=Number.From(EndDate),
    DateListNum={StartDateNum..EndDateNum},
    #"Converted to Table" = Table.FromList(DateListNum, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    ResultPre = Table.TransformColumnTypes(#"Converted to Table",{{"Column1", type date}}),
    Result= Table.RenameColumns(ResultPre,{{"Column1", "Date"}})
in
    Result
