(OriginalText as text,optional Keepdot as logical, optional Keepspace as logical) as text=>
let
    // MCode created by ThepExcel.com
    KeepdotFix=if Keepdot is null then true else Keepdot,
    KeepspaceFix=if Keepspace is null then true else Keepspace,
    ListToKeep={},
    ListToKeep2=if KeepdotFix then ListToKeep&{"."} else ListToKeep,
    ListToKeep3=if KeepspaceFix then ListToKeep2&{" "} else ListToKeep2,    
    Result = Text.Select(OriginalText,{"0".."9"}&ListToKeep3) 
    
in
    Result
