( optional vRootPath , optional vTopic ,optional vKeyWord , optional vFx )=>
let #"001" = ( vTextList , vKeyWordList )=> 
let #"002" = vTextList, 
#"003" = vKeyWordList, 
#"004" = ( txt , word )=> List.Select( List.Transform( Text.Split( txt , word ) , (word)=>Text.Trim(word) ) , (x)=>x<>"" ), 
#"005" = List.Transform( #"003", ( byWord )=> ( txt )=> #"004"(txt , byWord ) ), 
#"006" = List.Accumulate( #"005" , #"002" , ( txtList , fx )=> List.Combine( List.Transform( txtList , (x)=> fx(x) ) ) ) in #"006", 
#"007" = ".PQ" , 
#"008" = if vRootPath <> null and vRootPath <> "" then vRootPath else let x = try Expression.Evaluate( "RootPath" , #shared ) in if x[HasError] then error "【参数:根路径】错误：未使用根路径参数" else x[Value],
#"009" = if vTopic = null then "" else vTopic , 
#"010" = if vKeyWord = null then "" else vKeyWord , 
#"011" = if vFx = null or vFx = "" then (x)=> if Text.StartsWith( Table.ColumnNames(x){0} , "Column" ) then Table.PromoteHeaders(x) else x else vFx ,
#"012" = let x =  try Folder.Files( #"008" ) in if x[HasError] then error "未能打开指定目录，请确保目录参数正确。 " else x[Value], 
#"013" = Table.SelectRows( #"012" , each not Text.StartsWith([Name], "~$")), 
#"014" = Table.SelectRows(#"013", each Text.Contains([Folder Path], #"009")), 
#"015" = if #"010" = "" or #"010" = null then ";[Sheet]Sheet1" & #"007" else let x = Table.SelectRows( #"014", each Text.Contains([Name],#"010" ) and Text.Contains (Text.Upper( [Name] ), #"007" ) ), y = if Table.RowCount( x ) = 0 then error "未找到该信息标识文件。" else x[Name]{0} in y, 
#"016" = Text.Trim( Text.Upper( #"015" ) ), 
#"017" = #"001"( {Text.Split( #"016" , #"007" ){0} }, {";", "；"} ), 
#"018" = List.Select( #"017" , (item)=>Text.Contains( item , "TABLE]" ) or Text.Contains( item, "TABLE】" ) ), 
#"019" = #"001"( #"001"( #"018" , {"," , "，"} ) , { "[TABLE]" , "【TABLE】" , "【TABLE]" , "[TABLE】" } ), 
#"020" = List.Select( #"017" , ( item )=> Text.Contains( item , "SHEET]" ) or Text.Contains( item , "SHEET】" ) ), 
#"021" = #"001"( #"001"( #"020" ,{",", "，"} ) , { "[SHEET]" , "【SHEET】" , "【SHEET]" , "[SHEET】" } ) , 
#"022" = List.Select( #"017" , ( item )=> Text.Contains( item , "CSV]" ) or Text.Contains( item , "CSV】" ) ),
#"023" = #"001"(#"001"( #"022" , {",","，"} ) , { "[CSV]" , "【CSV】" , "【CSV]" , "[CSV" } ), 
#"024" = Table.SelectRows( #"014" , each Text.Upper( [Extension] ) = ".XLSX" ), 
#"025" = Table.SelectRows(#"014" , each Text.Upper( [Extension] ) = ".CSV" ), 
#"026" = Table.Buffer( Table.AddColumn( #"024" , "Data" , each Excel.Workbook( [Content] , false ) ) ) , 
#"027" = Table.Buffer( Table.AddColumn( #"025" , "Data", each Csv.Document( [Content] ) ) ), 
#"028" = Table.SelectColumns(#"027",{"Name", "Data"}), 
#"029" = Table.RenameColumns(#"028",{{"Data", "Item.Data"}}), 
#"030" = Table.AddColumn(#"029","Item.Name", each [Name] ), 
#"031" = Table.AddColumn(#"030", "Item.KIND", each "Csv"), 
#"032" = if Table.RowCount( #"028" ) > 0 then Table.ColumnNames( #"031"{0}[Item.Data] ) else error "没有发现CSV文件。", 
#"033" = Table.ExpandTableColumn( #"026" , "Data",{"Name", "Data", "Kind"}, {"Item.Name", "Item.Data", "Item.Kind"}), 
#"034" = Table.SelectRows( #"033" , each [Item.Kind] = "Table" and List.Contains( #"019" , Text.Upper( [Item.Name] ) ) ), 
#"035" = Table.SelectRows( #"033" , each [Item.Kind] = "Sheet" and List.Contains( #"021" , Text.Upper( [Item.Name] ) ) ), 
#"036" = Table.SelectColumns(#"034",{ "Name", "Item.Kind" , "Item.Name", "Item.Data" }), 
#"037" = Table.SelectColumns(#"035",{ "Name", "Item.Kind" , "Item.Name", "Item.Data" }), 
#"038" = Table.Combine( { #"036" , #"037" , #"031" } ), 
#"039" = Table.RenameColumns(#"038",{{"Name", "数据来源"},{"Item.Kind","来源类型"},{"Item.Name","来源名称"}}), 
#"040" = Table.TransformColumns( #"039",{ { "Item.Data", #"011" } } ), 
#"041" = if Table.RowCount( #"040" ) > 0 then "Y" else "N" , 
#"042" = if #"041" = "N" then error "未发现数据。" else List.Buffer( Table.ColumnNames( #"040"{0}[Item.Data] ) ), 
#"043" = if #"041" = "N" then error "未发现数据。" else Table.ExpandTableColumn( #"040", "Item.Data", #"042" , #"042" ), 
#"044" = #"043" in #"044"



-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------



( optional vRootPath , optional vTopic ,optional vKeyWord , optional vFx )=>let #"001" = (vTextList , vKeyWordList )=> let #"002" = vTextList, #"003" = vKeyWordList, #"004" = ( txt , word )=> List.Select( List.Transform( Text.Split( txt , word ) , (word)=>Text.Trim(word) ) , (x)=>x<>"" ), #"005" = List.Transform( #"003", ( byWord )=> 
( txt )=> #"004"(txt , byWord ) ), #"006" = List.Accumulate( #"005" , #"002" , ( txtList , fx )=> List.Combine( List.Transform( txtList , (x)=> fx(x) ) ) ) in #"006", #"007" = "PQ" , #"008" = if vRootPath <> null and vRootPath <> "" then vRootPath else let x = try Expression.Evaluate( "RootPath" , #shared ) in if x[HasError] then error "【参数:根路径】错误：未使用根路径参数" else x[Value],#"009" = if vTopic = null then "" else vTopic , #"010" = if vKeyWord = null then "" else vKeyWord , #"011" = if vFx = null or vFx = "" then (x)=> if Text.StartsWith( Table.ColumnNames(x){0} , "Column" ) then Table.PromoteHeaders(x) else x else vFx , #"012" = try Folder.Files( #"008" ) catch (e)=>error "未能打开指定目录，请确保目录参数正确。 ", #"013" = Table.SelectRows( #"012" , each not Text.StartsWith([Name], "~$")), #"014" = Table.SelectRows(#"013", each Text.Contains([Folder Path], #"009")), #"015" = if #"010" = "" or #"010" = null then ";[Sheet]Sheet1" & #"007" else let x = Table.SelectRows( #"014", each 
Text.Contains([Name],#"010" ) and Text.Contains( Text.Upper( [Name] ), #"007" ) ), y = if Table.RowCount( x ) = 0 then error "未找到该信息标识文件。" else x[Name]{0} in
y, #"016" = Text.Trim( Text.Upper( #"015" ) ), #"017" = #"001"( { Text.Split( #"016" , #"007" ){0} }, {";", "；"} ), #"018" = List.Select( #"017" , (item)=>
Text.Contains( item , "TABLE]" ) or Text.Contains( item, "TABLE】" ) ), #"019" = #"001"( #"001"( #"018" , {"," , "，"} ) , { "[TABLE]" , "【TABLE】" , "【TABLE]" ,
"[TABLE】" } ), #"020" = List.Select( #"017" , ( item )=> Text.Contains( item , "SHEET]" ) or Text.Contains( item , "SHEET】" ) ), #"021" = #"001"( #"001"( #"020" ,{",",
"，"} ) , { "[SHEET]" , "【SHEET】" , "【SHEET]" , "[SHEET】" } ) , #"022" = List.Select( #"017" , ( item )=> Text.Contains( item , "CSV]" ) or Text.Contains( item , "CSV】" ) ),
#"023" = #"001"(#"001"( #"022" , {",","，"} ) , { "[CSV]" , "【CSV】" , "【CSV]" , "[CSV" } ), #"024" = Table.SelectRows( #"014" , each Text.Upper( [Extension] ) = 
".XLSX" ), #"025" = Table.SelectRows(#"014" , each Text.Upper( [Extension] ) = ".CSV" ), #"026" = Table.Buffer( Table.AddColumn( #"024" , "Data" , each
Excel.Workbook( [Content] , false ) ) ) , #"027" = Table.Buffer( Table.AddColumn( #"025" , "Data", each Csv.Document( [Content] ) ) ), #"028" =
Table.SelectColumns(#"027",{"Name", "Data"}), #"029" = Table.RenameColumns(#"028",{{"Data", "Item.Data"}}), #"030" = Table.AddColumn(#"029","Item.Name", each
[Name] ), #"031" = Table.AddColumn(#"030", "Item.KIND", each "Csv"), #"032" = if Table.RowCount( #"028" ) > 0 then Table.ColumnNames( #"031"{0}[Item.Data] ) else error "没有发现CSV文件。", #"033" = Table.ExpandTableColumn( #"026" , "Data",{"Name", "Data", "Kind"}, {"Item.Name", "Item.Data", "Item.Kind"}), #"034" =
Table.SelectRows( #"033" , each [Item.Kind] = "Table" and List.Contains( #"019" , Text.Upper( [Item.Name] ) ) ), #"035" = Table.SelectRows( #"033" , each [Item.Kind] = 
"Sheet" and List.Contains( #"021" , Text.Upper( [Item.Name] ) ) ), #"036" = Table.SelectColumns(#"034",{ "Name", "Item.Kind" , "Item.Name", "Item.Data" }), #"037" =
Table.SelectColumns(#"035",{ "Name", "Item.Kind" , "Item.Name", "Item.Data" }), #"038" = Table.Combine( { #"036" , #"037" , #"031" } ), #"039" = 
Table.RenameColumns(#"038",{{"Name", "数据来源"},{"Item.Kind","来源类型"},{"Item.Name","来源名称"}}), #"040" = Table.TransformColumns( #"039",{ { "Item.Data", 
#"011" } } ), #"041" = if Table.RowCount( #"040" ) > 0 then "Y" else "N" , #"042" = if #"041" = "N" then error "未发现数据。" else
List.Buffer( Table.ColumnNames( #"040"{0}[Item.Data] ) ), #"043" = if #"041" = "N" then error "未发现数据。" else Table.ExpandTableColumn( #"040", "Item.Data", #"042" , #"042" ), #"044" = #"043" in #"044"