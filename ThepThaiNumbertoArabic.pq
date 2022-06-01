(OriginalText as text) as text =>
let
    //OriginalText = "แมว ๓๘ ตัว ราคา ๕๔๓๖ บาท",
    TextList=Text.ToList(OriginalText),
    ZipReplace=List.Zip({{"๐".."๙"},{"0".."9"}}),
    Result=List.ReplaceMatchingItems(TextList,ZipReplace,Comparer.OrdinalIgnoreCase),
    ResultCombine=Text.Combine(Result)
in
    ResultCombine
