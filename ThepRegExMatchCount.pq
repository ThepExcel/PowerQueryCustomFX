(OriginalText as text,RegExPattern as text, optional RegExMode as text) as number=>
 let 
     // MCode created by ThepExcel.com
    MyTextFix=Text.Replace(OriginalText,"\","\\"),
    MyTextFix2=Text.Replace(Text.Replace(MyTextFix,"#(lf)","\n"),"#(cr)","\r"),
    //convert \ to use in Javascript
    re= if RegExMode is null then "/"&RegExPattern&"/g" else "/"&RegExPattern&"/g"&RegExMode,
    // RegExMode is "g", "i" or "gi"
    //default mode is Single Match and case sensitive
    getWeb=Web.Page(
    "<meta http-equiv='X-UA-Compatible' content='IE=edge'>
    <script>
    var MyText='"&MyTextFix2&"';
    var MyMatches=MyText.match("&re&");
    document.write(MyMatches.length);
    </script>"),
    getWeb2=getWeb[Data]{0}[Children]{0}[Children]{1}[Text],
    finalResult=try Number.From(Text.Combine(getWeb2)) otherwise 0
    
in  
    finalResult
