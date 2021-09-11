(TableName as table,optional ReplaceWith as text) as table=>
let
    // MCode created by ThepExcel.com
    OriginalColumnName= Table.ColumnNames(TableName),
    NumCol=List.Count(OriginalColumnName),
    RepeatReplace=List.Repeat({ReplaceWith}, NumCol),
    zipList=List.Zip(   {  OriginalColumnName ,RepeatReplace  }  ), 
    Source =Table.ReplaceErrorValues(TableName,zipList )
in
    Source
