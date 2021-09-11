(StartDateText as text,EndDateText as text) as table=>
let
    // MCode created by ThepExcel.com
    DatefromText=(DateText as text,DateFormat as text,optional offsetYear as number)as date=>
    let
        offsetYearNew=if offsetYear is null then 0 else offsetYear,
        DateText2=Text.Select(Text.From(DateText),{"0".."9"}) ,
        YearNum=
        if DateFormat="yyyymmdd" then
            Number.From(Text.Start(DateText2,4))+offsetYearNew
        else
            Number.From(Text.End(DateText2,4))+offsetYearNew,
        MonthNum=
        if DateFormat="yyyymmdd" then
            Number.From(Text.Middle(DateText2,4,2))
        else if DateFormat="ddmmyyyy" then
            Number.From(Text.Middle(DateText2,2,2))
        else //mmddyyyy
            Number.From(Text.Start(DateText2,2))
        ,
        DayNum=
        if DateFormat="yyyymmdd" then
            Number.From(Text.End(DateText2,2))
        else if DateFormat="ddmmyyyy" then
            Number.From(Text.Start(DateText2,2))
        else //mmddyyyy
            Number.From(Text.Middle(DateText2,2,2))    
            ,
        DateResult=#date(YearNum,MonthNum,DayNum)   
    in
        DateResult
    ,
    //

    StartDate = DatefromText(StartDateText,"yyyymmdd"),
    EndDate=DatefromText(EndDateText,"yyyymmdd"),
    StartDateNum=Number.From(StartDate),
    EndDateNum=Number.From(EndDate),
    DateListNum={StartDateNum..EndDateNum},
    #"Converted to Table" = Table.FromList(DateListNum, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    ResultPre = Table.TransformColumnTypes(#"Converted to Table",{{"Column1", type date}}),
    Result= Table.RenameColumns(ResultPre,{{"Column1", "Date"}})
in
    Result
