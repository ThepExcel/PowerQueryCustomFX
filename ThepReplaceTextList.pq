(OriginalText as text,OldTextList as list,NewTextList as list) as text =>
let
    // MCode created by ThepExcel.com
    ItemCount=List.Count(OldTextList),
    ListGen={0..ItemCount-1},
    test=List.Accumulate(ListGen,OriginalText,(state,current)=>Text.Replace(state,OldTextList{current},NewTextList{current}))
    //test=Text.Replace(Source,"se","12")
in
    test