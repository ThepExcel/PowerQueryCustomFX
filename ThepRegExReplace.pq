(OriginalText as text,RegExPattern as text,NewText as text,optional RegExMode as text) as any=>
 let
     // MCode created by ThepExcel.com 
    MyTextFix=Text.Replace(OriginalText,"\","\\"),
    MyTextFix2=Text.Replace(Text.Replace(MyTextFix,"#(lf)","\n"),"#(cr)","\r"),
    NewTextFix=Text.Replace(NewText,"\","\\"),
    re= if RegExMode is null then "/"&RegExPattern&"/" else "/"&RegExPattern&"/"&RegExMode,
    // RegExMode is "g", "i" or "gi"
    //default mode is Single Match and case sensitive
    getWeb=Web.Page(
    "<meta http-equiv='X-UA-Compatible' content='IE=edge'>
    <script>
    var MyText='"&MyTextFix2&"';
    var MyNewText='"&NewTextFix&"';
    var MyMatches=MyText.replace("&re&",MyNewText);
    var MyResult=JSON.stringify(MyMatches);
    document.write(MyResult);
    </script>"),
    getWeb2=getWeb[Data]{0}[Children]{0}[Children]{1}[Text],
    finalResult=Json.Document(Text.Combine(getWeb2))
in  
    finalResult
