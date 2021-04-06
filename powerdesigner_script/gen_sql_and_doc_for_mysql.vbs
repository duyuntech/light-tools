'-------------------------------------------------------------------
' 描述(Discription)	: 模型生成脚本工具(PDM Build SQL Script Tool)
' 作者(Author)		: Kevin Penn(kevin.penn@outlook.com)
' 日期(Date)		: 2015 / 03 / 28 
' 版本(Version) 	: 1.0.0.0
' ------------------------------------------------------------------
' 变更 : 添加生成建表语句指定编码语句
' 原因 : MySQL插入中文字符时报错
' 描述 : MySQL字符要统一,包括服务端、客户端、数据库、表、连接字符串
'		 编码统一
' 日期 : 2015 / 09 / 13
'-------------------------------------------------------------------
' 变更 : 添加约束SQL生成
' 原因 : 业务中有需要唯一约束的业务场景
' 描述 : *
' 日期 : 2016 / 04 / 20
'-------------------------------------------------------------------
Option Explicit
InteractiveMode = im_Batch

Const DEBUG_MODEL = True				                        ' 是否开启调试模式
Const Capslock = False				            	            ' 是否开启大写模式
Const DIR_TARGET  = "D:\Gendoc\"		                    ' 脚本输出目录
Const TABLE_ENCODE = "UTF8"				                      ' 输出表的编码
Const DOC_TIME = True				            	              ' 是否输出文档日期

' 字段类型支持列表
Dim DType,idx
DType = Array("INT","BIGINT","CHAR")

'For idx=0 To UBound(DType)-LBound(DType)
'    MsgBox DType(idx)
'Next


Dim workspace, fileStream, commandShell, fileTarger, htmlFileTarger, curHtmlFileName
Set workspace = ActiveWorkspace
Set fileStream = createobject("Scripting.FileSystemObject")	
set commandShell = createobject("WScript.Shell") 


If (workspace Is Nothing) Then
	MsgBox "工作空间中没有活动的视图(There is no Active Diagram)!"
Else
		
	Dim children, folder
	Set children = workspace.Children	

	DelFiles(DIR_TARGET)
	MakeDir(DIR_TARGET)
   
   	Debug(workspace)
	For Each folder In children
		Debug("Folder : " + folder)
		If(folder.ModelObject.IsKindOf(PdPDM.Cls_Model)) Then
			Debug("Folder name :" + folder.Name)
			ListModel folder.ModelObject
			GenDocs folder.ModelObject
		End If
	Next
End If	


' 扫描PDM文件中的物理模型文件
' Scan PDM file Logic model 
Function ListModel(model)
	if(model Is Nothing) Then
		MsgBox 'Nothing Model'
	Else
    	Dim curColumn, curTable    
		For Each curTable In model.Tables : ProcessTable(curTable) : Next
	End If
	commandShell.run """" + DIR_TARGET + """"
End Function

' Generate Docs
Function GenDocs(model)
	if(model Is Nothing) Then
		MsgBox 'Nothing Model'
	Else
    	
    	Dim curColumn, eachColumn,curTable, curCommontHtml
		
		curHtmlFileName = model.Name + "_docs"

		WriteHtml curHtmlFileName, "<html><head><title>" + curHtmlFileName + " - 数据库注释</title><style>body {margin:0 auto;width:90%;}a {text-decoration:none;font-size:14px;} table.hovertable {font-family: verdana,arial,sans-serif;font-size:14px;color:#333333;border-width: 1px;border-color: #999999;border-collapse: collapse;width:100%;}._tb_head th {background-color:#D6EEF0;border-width: 1px;padding: 8px;border-style: solid;border-color: #a9c6c9; color: #555} table.hovertable td {border-width: 1px;padding: 8px;border-style: solid;border-color: #a9c6c9;} .col_name,.col_type {width:200px} ._tb_name {text-align: left; color: #000;padding:10px;} p {font-size: 10px;color: #ccc;margin-top: 5px;}</style></head><body>"
		
		If(DOC_TIME) Then
			WriteHtml curHtmlFileName, "<p>修改日期 : " + FormatTime(Now,4) + "</p>"
		End If	
		
		' 生成索引
		WriteHtml curHtmlFileName, "<ol>"
		For Each curTable In model.Tables 
			curCommontHtml = "<li><a href='#" + curTable.Code + "'>" + curTable.Code + "/" + curTable.Comment + "</a></li>"
			WriteHtml curHtmlFileName, curCommontHtml
		Next
		WriteHtml curHtmlFileName, "</ol>"
		
		For Each curTable In model.Tables 
			curCommontHtml = "<a name='" + curTable.Code + "' id='" + curTable.Code + "'></a>"
			curCommontHtml = curCommontHtml + "<table class='hovertable'><tr><th colspan='3' class='_tb_name'>" + curTable.Comment + " / " + curTable.Code  + "</th></tr><tr class='_tb_head'><th class='col_name'>字段名称</th><th class='col_type'>字段类型</th><th>字段说明</th></tr>"
			'loop Columns
			For Each eachColumn In curTable.Columns 
				curCommontHtml = curCommontHtml + "<tr><td>" + UCase(eachColumn.Code) + "</td><td>" + UCase(eachColumn.DataType) + "</td><td>" + eachColumn.Comment + "</td></tr>"		
			Next
			curCommontHtml = curCommontHtml + "</table><br />"
			WriteHtml curHtmlFileName, curCommontHtml
		Next
		
		WriteHtml curHtmlFileName, "</body></html>"
		
	End If
End Function


' 处理表并生成建表语句
' Proccess table and build statement 
Function ProcessTable(argTable)
		Dim curTabSql,eachColumn
		Debug "Table code : " + argTable.Code + " : Table name : " + argTable.Name	
		
		curTabSql = curTabSql + "DROP TABLE IF EXISTS " + argTable.Code + ";" + VBCRLF
		curTabSql = curTabSql + "CREATE TABLE " + argTable.Code + VBCRLF
		curTabSql = curTabSql + "(" + VBCRLF
		'msgbox CStr(UBound(curTable.Columns))
		
		'loop Columns
		For Each eachColumn In argTable.Columns 
		
			curTabSql = curTabSql + Left(Space(3) + UCase(eachColumn.Code) + Space(30), 30) + Space(4) + UCase(eachColumn.DataType)
			' Is Not Null
			If((eachColumn.NullStatus = "not null") And (Not eachColumn.Primary)) Then
				curTabSql = curTabSql + " NOT NULL," + VBCRLF	
			Else
				' Is Default
				If(eachColumn.DefaultValueDisplayed = "") Then
					curTabSql = curTabSql + "," + VBCRLF
				Else
					curTabSql = curTabSql + " DEFAULT '" + eachColumn.DefaultValueDisplayed + "'," + VBCRLF
				End If	
			End If
			
			Debug("Column DefaultValue:" + eachColumn.DefaultValueDisplayed)			
		Next
		
		curTabSql = Left(curTabSql, InStrRev(curTabSql,",") - 1) + VBCRLF
		curTabSql = curTabSql + ") ENGINE=INNODB,DEFAULT CHARSET " + TABLE_ENCODE + ";" + VBCRLF
		curTabSql = curTabSql + ProccessPk(argTable) + VBCRLF
		curTabSql = curTabSql + ProccessUnion(argTable) + VBCRLF
		curTabSql = curTabSql + ProccessIndex(argTable) + VBCRLF + VBCRLF
		WriteSql argTable, curTabSql
		
		Debug(curTabSql)
		
End Function

' 处理主键
' Proccess Primary Key
Function ProccessPk(argTable)
	ProccessPk = ""
	Dim pk,column,lastPK,lastPkType,lastPkName
	Set pk = argTable.PrimaryKey
	If(pk Is Nothing) Then
		output "Warning: " + argTable.Code + "[无主键]"
		Exit Function
   	Else
   		output "ProccessPk: " + pk.Code
   		
		ProccessPk = ProccessPk + "ALTER TABLE " + argTable.Code + " ADD CONSTRAINT " + pk.Code + " PRIMARY KEY ("
		
		For Each column In pk.Columns

			lastPkName = column.Code
			lastPkType = column.DataType

			ProccessPk = ProccessPk + column.Code + ","
			' 如果为自增长则记录
			If(column.Identity) Then
				lastPK = column.Code
			End If
		Next
		ProccessPk = Left(ProccessPk, InStrRev(ProccessPk,",") - 1)
		ProccessPk = ProccessPk + ");"
   		' 主键自增长
		If(lastPK <> "") Then
			ProccessPk = ProccessPk + VBCRLF + "ALTER TABLE " + argTable.Code + " CHANGE " + lastPkName + " " + lastPkName + " " + lastPkType + " NOT NULL AUTO_INCREMENT;"
			ProccessPk = ProccessPk + VBCRLF + "ALTER TABLE " + argTable.Code + " AUTO_INCREMENT = 10000;"
		End If
	End If
End Function


' 处理唯一约束主键
' Proccess Union
' ALTER TABLE `HCS_MEDICINE_ORG` ADD unique(`MEDICINE_ID`,`ORG_ID`,`ORG_MEDICINE_ID`);
Function ProccessUnion(argTable)
	ProccessUnion = ""
	Dim keys,key,column
	Set keys = argTable.Keys
	If(keys Is Nothing) Then
		'output "Warning: " + argTable.Code + "[无主键]"
		Exit Function
   	Else 
		For Each key In keys
			' 唯一约束
			if(key.Primary = false) Then
				ProccessUnion = "ALTER TABLE `" + argTable.Code + "` ADD UNIQUE ("
				For Each column In key.Columns
					ProccessUnion = ProccessUnion + "`" + column.Code + "`,"
				Next
				ProccessUnion = Left(ProccessUnion, InStrRev(ProccessUnion,",") - 1)
				ProccessUnion = ProccessUnion + ");"
				' MsgBox ProccessUnion
			End If
		Next   		
	End If
End Function

' 处理表的索引
' Proccess Index
Function ProccessIndex(argTable)
	ProccessIndex = ""

	Dim curIndex,column
	For Each curIndex In argTable.Indexes
	
		Debug "Index:" + curIndex.Code
		
		ProccessIndex = ProccessIndex + "CREATE" + Space(1)
		
		If(curIndex.Unique = true) Then
			Debug "Index:Unique"
			ProccessIndex = ProccessIndex + "UNIQUE" + Space(1)
		End If
		
		ProccessIndex = ProccessIndex + "INDEX " + curIndex.Code + Space(1) + "ON" + Space(1) + argTable.Code + Space(1) + "("
		
		For Each column In curIndex.IndexColumns
			ProccessIndex = ProccessIndex + column.Code + ","
		Next
		ProccessIndex = Left(ProccessIndex, InStrRev(ProccessIndex,",") - 1)
		ProccessIndex = ProccessIndex + ");" + VBCRLF
	Next
End Function

' 检查目标文件是否存在
' Vary target file
Function VaryFile(argTable, postfix)
	If(argTable Is Nothing) Then
		Set fileTarger = fileStream.OpenTextFile(DIR_TARGET + "Call_Error" + postfix, 8, true)	
	Else
		Set fileTarger = fileStream.OpenTextFile(DIR_TARGET + Owner(argTable) + postfix, 8, true)		
	End If	

End Function

' Open File
Function OpenFile(filename, postfix)
	' 第四个参数为编码
	' -2	以系统默认格式打开文件。
	' -1	以 Unicode 格式打开文件。
	' 0		以 ASCII 格式打开文件。
	Set htmlFileTarger = fileStream.OpenTextFile(DIR_TARGET + filename + postfix, 8, true, -1)		
End Function

' 写入SQL文件
' Write Sql to file
Function WriteSql(argTable, sqlStatement)
	VaryFile argTable , ".sql"
	If(Capslock) Then
		fileTarger.Write(UCase(sqlStatement))
	Else
		fileTarger.Write(LCase(sqlStatement))
	End If
	fileTarger.Close()
End Function

' 写入Html文件
' Write Html to file
Function WriteHtml(filename, content)
	OpenFile filename , ".html"
	' 内容中含有 Unicode 格式的字符报错
	If(Capslock) Then
		htmlFileTarger.Write(UCase(Replace(content, chrw(8226), "")))
	Else
		htmlFileTarger.Write(LCase(Replace(content, chrw(8226), "")))
	End If
	
	htmlFileTarger.Close()
End Function

' 获取当前表的所有者
' Get owner for current table 
Function Owner(argTable)
	If(argTable.Owner Is Nothing) Then
		Owner = "Unkonw"
	Else
		Owner = argTable.Owner.Code
	End If
End Function

' 格式化时间
' Format Datatime
Function FormatTime(argTime, argFlag)
	Dim y, m, d, h, mi, s
	FormatTime = "19700101000000"
	If IsDate(argTime) = False Then Exit Function
	
	y  = cstr( year(argTime))
	m  = cstr( month(argTime))
	d  = cstr( day(argTime))
	h  = cstr( hour(argTime))
	mi = cstr( minute(argTime))
	s  = cstr( second(argTime))
	
	If len(m)  = 1 Then m  = "0" & m	
	If len(d)  = 1 Then d  = "0" & d	
	If len(h)  = 1 Then h  = "0" & h	
	If len(mi) = 1 Then mi = "0" & mi	
	If len(s)  = 1 Then s  = "0" & s
		
	Select Case argFlag
		Case 1
			' 20150101121001
			FormatTime = y & m & d & h & mi & s
		Case 2
			' 2015-01-01
			FormatTime = y & "-" & m & "-" & d
		Case 3
			' 12:10:10
			FormatTime = h & ":" & mi & ":" & s
		Case 4
			' 2015年04月04日18时55分17秒
			FormatTime = y & "年" & m & "月" & d & "日 " & h & "时" & mi & "分" & s & "秒"
		Case 5
			' 20140101
			FormatTime = y & m & d
	End Select
End Function

' 创建文件夹
' Create folder
Function MakeDir(path)
	If fileStream.FolderExists(path) Then
		Exit Function
	End If
	If Not fileStream.FolderExists(fileStream.GetParentFolderName(path)) Then
		create fileStream,fileStream.GetParentFolderName(path)  
	End If
	fileStream.CreateFolder(path)
End Function

' 从路径中删除所有文件
' Delete files for path
Function DelFiles(dirPath)
	Dim folder,subFolder,file
	Set folder = fileStream.GetFolder(dirPath)
	
	For Each file In folder.Files : file.delete : Next
	
	For Each subFolder in folder.SubFolders : fileStream.DeleteFolder subFolder : Next
End Function

' 打印调试信息
' Print debug message
Function Debug(message)
	If(DEBUG_MODEL) Then
		'output Left("***DEBUG-INFO:" + FormatTime(Now,4) + "*******************************************************************",80)
		output message
		'output Left("***INFO - END:" + FormatTime(Now,4) + "*******************************************************************",80)
	End If
End Function
