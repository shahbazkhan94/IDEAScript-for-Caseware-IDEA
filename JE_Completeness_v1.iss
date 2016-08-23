Dim ListBox1$() AS string
Dim ListBox2$() AS string
Dim ListBox3$() AS string
Dim ListBox4$() AS string
Dim ListBox5$() AS string
Dim ListBox6$() AS string

Begin Dialog dlgMenu 49,10,339,215,"Journal Entries Completeness Task", .Displayit
  Text 58,16,100,28, "Text", .txtfname1
  Text 58,65,100,24, "Text", .txtfname2
  Text 58,113,100,27, "Text", .txtfname3
  Text 171,15,50,14, "TB_OP_Bal Field", .Text2
  Text 171,63,50,10, "TB_CL_Bal Field", .Text2
  Text 171,112,50,14, "JE_DR Field", .Text2
  Text 171,129,50,10, "JE_CR Field", .Text2
  DropListBox 231,15,84,11, ListBox1$(), .DropListBox1
  DropListBox 231,63,84,11, ListBox2$(), .DropListBox3
  DropListBox 231,112,84,10, ListBox3$(), .DropListBox5
  DropListBox 231,130,84,10, ListBox3$(), .DropListBox6
  Text 12,172,40,14, "New Filename", .Text2
  TextBox 61,172,106,12, .txtNewFilename
  PushButton 12,16,40,14, "&Select file", .PushButton1
  PushButton 12,64,40,14, "Selec&t file", .PushButton2
  PushButton 12,113,40,14, "Sel&ect file", .PushButton3
  OKButton 198,172,40,14, "O&k", .OKButton1
  CancelButton 269,172,40,14, "Can&cel", .CancelButton1
  Text 171,33,50,14, "TB_OP_Bal Match Field", .Text2
  DropListBox 231,33,84,11, ListBox4$(), .DropListBox2
  Text 171,78,50,14, "TB_CL_Bal Match Field", .Text2
  DropListBox 231,78,84,11, ListBox5$(), .DropListBox4
  Text 171,146,50,14, "JE db Match Field", .Text2
  DropListBox 231,147,84,11, ListBox6$(), .DropListBox7
  GroupBox 7,4,315,45, "Trial Balance for Openning Balance", .GroupBox1
  GroupBox 7,100,315,62, "Journal Entries file", .GroupBox2
  GroupBox 7,52,315,42, "Trial Balance for Closing Balance", .GroupBox2
End Dialog
'****************************************************************************************************************
'* Script:		JE_Completeness.iss
'* By:		Shahbaz Khan
'* Version:	1.0
'* Date:		March 25, 2016
'* Purpose:	To ensure completeness of JEs by trial balance movement 
'***************************************************************************************************************
Option Explicit

Dim fname1 As String  'Trail Balance for openning balance
Dim fname2 As String  'Trial Balance for closing balance
Dim fname3 As String 'JEs file
Dim fname4 As String ' Summarized JEs file
Dim fname5 As String ' Joined all files
Dim amtfield1 As String 	' for tb OP bal
Dim amtfield2 As String 	' for tb OP match 
Dim amtfield3 As String	' for tb CL bal
Dim amtfield4 As String	' for tb CL match
Dim amtfield5 As String	' for JE Dr bal
Dim amtfield6 As String	' for JE Cr bal
Dim amtfield7 As String	' for JE match
Dim newFilename As String	'new filename'
Dim working_directory As String
Dim exitScript As Boolean
Dim returnsubmain As Boolean

Sub Main
	working_directory = Client.WorkingDirectory
	Call menu()
	If Not exitScript Then
		Call Summarization()
		Call RelateDatabase()
		Call AppendField()
		Call AppendField1()
	End If
	client.refreshFileExplorer
End Sub

Function menu()
	'Local variable definition
	Dim dlg As dlgMenu
	Dim button As Integer
	Dim filebar As Object
	Dim exitDialog As Boolean
	Dim source As Object
	Dim table As Object
	Dim fields As Integer
	Dim i, j, k As Integer
	Dim field As Object
	
	'Looping for dialog display
	Do
	
		button = Dialog(dlg)
		
		Select Case button
			Case -1 'ok button 
				If dlg.DropListBox1 > -1 Then
					amtfield1 = ListBox1$(dlg.DropListBox1)
				Else
					amtfield1 = ""
				End If
								
				If dlg.DropListBox2 > -1 Then
					amtfield2 = ListBox4$(dlg.DropListBox2)		' ListBox4 for all fields types for match key OP TB
				Else
					amtfield2 = ""
				End If
			
				If dlg.DropListBox3 > -1 Then
					amtfield3 = ListBox2$(dlg.DropListBox3)
				Else
					amtfield3 = ""
				End If
				
				If dlg.DropListBox4 > -1 Then
					amtfield4 = ListBox5$(dlg.DropListBox4)		' ListBox5 for all fields types for match key CL TB
				Else
					amtfield4 = ""
				End If
				
				If dlg.DropListBox5 > -1 Then
					amtfield5 = ListBox3$(dlg.DropListBox5)
				Else
					amtfield5 = ""
				End If
				
				If dlg.DropListBox6 > -1 Then
					amtfield6 = ListBox3$(dlg.DropListBox6)
				Else
					amtfield6 = ""
				End If
				
				If dlg.DropListBox7 > -1 Then				' ListBox6 for all fields types for match key JE db
					amtfield7 = ListBox6$(dlg.DropListBox7)
				Else
					amtfield7 = ""
				End If
			
				newFilename = dlg.txtNewFilename
				if validatemenu() then exitDialog = TRUE
			Case 0 ' cancel button
				exitDialog = TRUE
				exitScript = TRUE
				
			Case 1 ' File for TB OP Balance
				Set filebar = CreateObject("ideaex.fileexplorer")
				filebar.displaydialog
				fname1 = filebar.selectedfile
				If fname1 <> "" Then
					Set source = client.opendatabase(fname1)
					Set table = source.tabledef
					fields = table.count
					ReDim ListBox1$(fields)
					ReDim ListBox4$(fields)
					j = 0
					k = 0
					For i = 1 To fields
						Set field = table.getfieldat(i)
						If field.isnumeric Then
							ListBox1$(j) = field.name
							j = j + 1
						End If
						ListBox4$(k) = field.name
						k = k +1
					Next i
					
				End If
				
				If j = 0 Then
					MsgBox "The file selected does not contain a numeric field", MB_ICONEXCLAMATION, "Error 5"
					fname1 = ""
				End If
			
			Case 2 ' File for TB CL Balance
				Set filebar = CreateObject("ideaex.fileexplorer")
				filebar.displaydialog
				fname2 = filebar.selectedfile
				If fname2 <> "" Then
					Set source = client.opendatabase(fname2)
					Set table = source.tabledef
					fields = table.count
					ReDim ListBox2$(fields)
					ReDim ListBox5$(fields)
					j = 0
					k = 0
					For i = 1 To fields
						Set field = table.getfieldat(i)
						If field.isnumeric Then
							ListBox2$(j) = field.name
							j = j + 1
						End If
						ListBox5$(k) = field.name
						k = k + 1
					Next i
					
				End If
				
				If j = 0 Then
					MsgBox "The file selected does not contain a numeric field", MB_ICONEXCLAMATION, "Error 5"
					fname2 = ""
				End If

			Case 3 ' File for JE DR CR Balance
				Set filebar = CreateObject("ideaex.fileexplorer")
				filebar.displaydialog
				fname3 = filebar.selectedfile
				If fname3 <> "" Then
					Set source = client.opendatabase(fname3)
					Set table = source.tabledef
					fields = table.count
					ReDim ListBox3$(fields)
					ReDim ListBox6$(fields)
					j = 0
					k = 0
					For i = 1 To fields
						Set field = table.getfieldat(i)
						If field.isnumeric Then
							ListBox3$(j) = field.name
							j = j + 1
						End If
						ListBox6$(k) = field.name
						k = k + 1
					Next i
					
				End If
				
				If j = 0 Then
					MsgBox "The file selected does not contain a numeric field", MB_ICONEXCLAMATION, "Error 5"
					fname3 = ""
				End If

								
		End Select
	Loop While exitDialog = FALSE
	
	'Clearing memory
	Set source = Nothing
	Set table = Nothing
	Set field = Nothing
End Function

Function validateMenu() As Boolean
	Dim db1 As database
	Dim db2 As database
	Dim db3 As database
	Dim table1 As table
	Dim table2 As table
	Dim table3 As table
	Dim field1 As Object
	Dim field2 As Object
	Dim field3 As Object
	
	validateMenu = TRUE
	'Error 1 - TB Op file selection check with relevant fields
	If fname1 = "" Then
		MsgBox "Please select a trial balance for openning balance.", MB_ICONEXCLAMATION, "Error 1"
		validateMenu = FALSE
	ElseIf fname2 = "" Then
		MsgBox "Please select a trial balance for closing balance.", MB_ICONEXCLAMATION, "Error 1"
		validateMenu = FALSE
	ElseIf fname3 = "" Then
		MsgBox "Please select a journal entries file.", MB_ICONEXCLAMATION, "Error 1"
		validateMenu = FALSE
	End If
	
	'Error 2 - amount fields and Error 3 - for match fields for joining database
	If amtfield1 = "" Then
		MsgBox "Please select proper openning balance field.", MB_ICONEXCLAMATION, "Error 2"
		validateMenu = FALSE
	ElseIf amtfield2 = "" Then 
		MsgBox "Please select match key field for openning trial balance database.", MB_ICONEXCLAMATION, "Error 3"
		validateMenu = FALSE
	ElseIf amtfield3 = "" Then 
		MsgBox "Please select proper openning balance field.", MB_ICONEXCLAMATION, "Error 2"
		validateMenu = FALSE
	ElseIf amtfield4 = "" Then 
		MsgBox "Please select match key field for closing trial balance database.", MB_ICONEXCLAMATION, "Error 3"
		validateMenu = FALSE
	ElseIf amtfield5 = "" Then 
		MsgBox "Please select proper debit balance field from journal entries file.", MB_ICONEXCLAMATION, "Error 2"
		validateMenu = FALSE
	ElseIf amtfield6 = "" Then 
		MsgBox "Please select proper credit balance field from journal entries file.", MB_ICONEXCLAMATION, "Error 2"
		validateMenu = FALSE
	ElseIf amtfield7 = "" Then 
		MsgBox "Please select match key field for Journal Entries database.", MB_ICONEXCLAMATION, "Error 3"
		validateMenu = FALSE
	End If 

	'Error 3 - newfilename error
	If newFilename = "" Then
		MsgBox "Please enter a new filename", MB_ICONEXCLAMATION, "Error 3"
		validateMenu = FALSE
	End If
	
	'Error 4 - checkfor special character in filename
	If checkForSpecialChar(newFilename, "\/:*?""<>[]|") Then
		MsgBox "Please do not use the following in your filename - \/:*?""<>[]|", MB_ICONEXCLAMATION, "Error 4"
		validateMenu = FALSE
	End If
	
	
	'Error 6 - to check matching key fields for the same type for joining
	Set db1 = Client.OpenDatabase(fname1)
	Set db2 = Client.OpenDatabase(fname2)
	Set db3 = Client.OpenDatabase(fname3)
	Set table1 = db1.TableDef
	Set table2 = db2.TableDef
	Set table3 = db3.TableDef
	Set field1 = table1.GetField(amtfield2)
	Set field2 = table2.GetField(amtfield4)
	Set field3 = table3.GetField(amtfield7)
	If field1.Type = field2.Type And field2.Type = field3.Type Then
		validateMenu = TRUE
		Client.CloseAll
	Else
		MsgBox "Field type of matching keys are not same. Please select same field type.", MB_ICONEXCLAMATION, "Error 6"
		validateMenu = FALSE
		Client.CloseAll
	End If 
	Set db1 = Nothing
	Set db2 = Nothing
	Set db3 = Nothing
	Set table1 = Nothing
	Set table2 = Nothing
	Set table3 = Nothing
	Set field1 = Nothing
	Set field2 = Nothing
	Set field3 = Nothing
	
End Function

'Function to display the dialog
Function Displayit(ControlID$, Action%, SuppValue%)
	If fname1 = "" Then
		DlgText "txtfname1", "No file selected"
	Else
		DlgText "txtfname1", getFileName(fname1,0)
	End If
	
	If fname2 = "" Then
		DlgText "txtfname2", "No file selected"
	Else
		DlgText "txtfname2", getFileName(fname2,0)
	End If
	
	If fname3 = "" Then
		DlgText "txtfname3", "No file selected"
	Else
		DlgText "txtfname3", getFileName(fname3,0)
	End If

End Function	

' Analysis: Summarization
Function Summarization
	Dim db As database
	Dim task As task
	Dim dbName As String
	
	Set db = Client.OpenDatabase(fname3)
	Set task = db.Summarization
	task.UseQuickSummarization = TRUE
	task.AddFieldToSummarize amtfield7
	task.AddFieldToTotal amtfield5
	task.AddFieldToTotal amtfield6
	dbName = client.UniqueFilename("Summarization by Account.IMD")
	Set fname4 = dbName
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function



' File: Visual Connector
Function RelateDatabase
	Dim db As database
	Dim task As task
	Dim dbName As String
	Dim id0 As Variant
	Dim id1 As Variant
	Dim id2 As Variant
	
	Set db = Client.OpenDatabase(fname2)
	Set task = db.VisualConnector
	Set id0 = task.AddDatabase(fname2)
	Set id1 = task.AddDatabase(fname1)
	Set id2 = task.AddDatabase(fname4)
	task.MasterDatabase = id0
	task.AppendDatabaseNames = FALSE
	task.IncludeAllPrimaryRecords = TRUE
	dbName = client.UniqueFilename(newFilename)
	Set fname5 = dbName
	task.AddRelation id0, amtfield4, id1, amtfield2
	task.AddRelation id1, amtfield2, id2, amtfield7
	task.IncludeAllFields
	task.CreateVirtualDatabase = False	
	task.OutputDatabaseName = dbName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set id0 = Nothing
	Set id1 = Nothing
	Set id2 = Nothing
	Exit Sub						
End Function

' Append Field : DERIVED_CLOSING
Function AppendField
	Dim db As database
	Dim task As task
	Dim dbName As String
	Dim field As Object
	Dim eqn As String
	
	Set db = Client.OpenDatabase(fname5)
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DERIVED_CLOSING"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	eqn = amtfield1 & "+" & amtfield5 &"_SUM" & "-" & amtfield6 &"_SUM"
	field.Equation = eqn
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set field = Nothing
End Function

'Append Field 1 : DIFFERENCE
Function AppendField1
	Dim db As database
	Dim task As task
	Dim dbName As String
	Dim field As Object
	Dim eqn1 As String
	
	Set db = Client.OpenDatabase(fname5)
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DIFFERENCE"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	eqn1 =  amtfield1 & "+" & amtfield5 &"_SUM" & "-" & amtfield6 &"_SUM" & "-" & amtfield3
	field.Equation = eqn1
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set field = Nothing
	db.ControlAmountField "DIFFERENCE"
	Client.CloseAll
	Client.OpenDatabase(fname5)
End Function

Function getFileName(temp_filename As String, temp_type As Boolean) '1 if get the name with any folder info, 0 if only the name
	Dim temp_length As Integer
	Dim temp_len_wd As Integer
	Dim temp_difference As Integer
	Dim temp_char As String
	Dim tempfilename As String
	
	If temp_type Then
		temp_len_wd  = Len(working_directory )  + 1'get the lenght of the working directory
		temp_length = Len(temp_filename) 'get the lenght of the file along with the working directory
		temp_difference = temp_length - temp_len_wd  + 1'get the lenght of just the filename
		getFileName = Mid(temp_filename, temp_len_wd, temp_difference)	
	Else
		temp_length  = Len(temp_filename )
		Do 
			temp_char = Mid(temp_filename, temp_length , 1)
			temp_length = temp_length  - 1 
			If temp_char <> "\" Then
				tempfilename = temp_char & tempfilename
			End If
		Loop Until temp_char = "\" Or temp_length = 0
		getFileName = tempfilename
	End If
End Function

Function checkForSpecialChar(temp_string As String, temp_list As String) As Boolean
	Dim strLen As Integer
	Dim tempChar As String
	Dim i As Integer
	Dim pos As Integer
	checkForSpecialChar = FALSE
	strlen = Len(temp_list)
	For i = 1 To strLen
		tempChar = Mid(temp_list, i, 1)
		pos = InStr(1, temp_string, tempChar)
		If pos > 0 Then
			checkForSpecialChar = TRUE
		End If
	Next i
End Function