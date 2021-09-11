(OriginalText as text,optional TrimChar as text) as text=>
let
    // MCode created by ThepExcel.com
    TrimCharFix = if TrimChar is null then " " else TrimChar,
    SplitStep = Text.Split(OriginalText,TrimCharFix),
    RemoveBlankStep = List.Select(SplitStep, each _ <> ""),
    CombineStep = Text.Combine(RemoveBlankStep,TrimCharFix)
in
CombineStep
