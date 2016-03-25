Dim ListBox1$() AS string
Dim ListBox2$() AS string
Dim ListBox3$() AS string

Begin Dialog dlgMenu 51,11,338,216,"Journal Entries Completeness Task", .Displayit
  Text 6,6,96,7, "Trial Balance for Openning Balance", .Text1
  Text 6,50,86,9, "Trial Balance for Closing Balance", .Text1
  Text 7,92,40,10, "Journal Entries", .Text1
  Text 54,16,100,28, "Text", .txtfname1
  Text 55,63,99,24, "Text", .txtfname2
  Text 54,105,100,27, "Text", .txtfname3
  Text 167,15,40,14, "TB_OP_Bal Field", .Text2
  Text 168,62,40,14, "TB_CL_Bal Field", .Text2
  Text 168,102,40,14, "JE_DR Field", .Text2
  Text 167,119,40,14, "JE_CR Field", .Text2
  DropListBox 227,16,84,11, ListBox1$(), .DropListBox1
  DropListBox 227,63,84,11, ListBox2$(), .DropListBox3
  DropListBox 227,105,84,10, ListBox3$(), .DropListBox5
  DropListBox 227,120,84,10, ListBox3$(), .DropListBox6
  Text 8,172,40,14, "New Filename", .Text2
  TextBox 57,171,106,12, .txtNewFilename
  PushButton 7,16,40,14, "Select file", .PushButton1
  PushButton 7,62,40,14, "Select file", .PushButton2
  PushButton 8,105,40,14, "Select file", .PushButton3
  OKButton 194,172,40,14, "OK", .OKButton1
  CancelButton 265,173,40,14, "Cancel", .CancelButton1
  Text 167,33,40,14, "TB_OP_Bal Match Field", .Text2
  DropListBox 227,34,84,11, ListBox1$(), .DropListBox2
  Text 167,76,40,14, "TB_CL_Bal Match Field", .Text2
  DropListBox 227,77,84,11, ListBox2$(), .DropListBox4
  Text 167,140,40,14, "JE db Match Field", .Text2
  DropListBox 227,141,84,11, ListBox3$(), .DropListBox7
End Dialog
'****************************************************************************************************************
'* Script:		JE_Completeness.iss
'* By:		Shahbaz
'* Version:	1.0
'* Date:		March 24, 2016
'* Purpose:	To ensure completeness of JEs by trial balance movement for millions of records
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
	Dim i, j As Integer
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
					amtfield2 = ListBox1$(dlg.DropListBox2)
				Else
					amtfield2 = ""
				End If
			
				If dlg.DropListBox3 > -1 Then
					amtfield3 = ListBox2$(dlg.DropListBox3)
				Else
					amtfield3 = ""
				End If
				
				If dlg.DropListBox4 > -1 Then
					amtfield4 = ListBox2$(dlg.DropListBox4)
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
				
				If dlg.DropListBox7 > -1 Then
					amtfield7 = ListBox3$(dlg.DropListBox7)
				Else
					amtfield7 = ""
				End If
			
				newFilename = dlg.txtNewFilename
				exitDialog = TRUE
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
					j = 0
					For i = 1 To fields
						Set field = table.getfieldat(i)
						If field.isnumeric Then
							ListBox1$(j) = field.name
							j = j + 1
						End If
					Next i
					
				End If
				
				If j = 0 Then
					MsgBox "The file selected does not contain a numeric field", MB_ICONEXCLAMATION, "Error"
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
					j = 0
					For i = 1 To fields
						Set field = table.getfieldat(i)
						If field.isnumeric Then
							ListBox2$(j) = field.name
							j = j + 1
						End If
					Next i
					
				End If
				
				If j = 0 Then
					MsgBox "The file selected does not contain a numeric field", MB_ICONEXCLAMATION, "Error"
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
					j = 0
					For i = 1 To fields
						Set field = table.getfieldat(i)
						If field.isnumeric Then
							ListBox3$(j) = field.name
							j = j + 1
						End If
					Next i
					
				End If
				
				If j = 0 Then
					MsgBox "The file selected does not contain a numeric field", MB_ICONEXCLAMATION, "Error"
					fname3 = ""
				End If

								
		End Select
	Loop While exitDialog = FALSE
	
	'Clearing memory
	Set source = Nothing
	Set table = Nothing
	Set field = Nothing
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
	task.AddRelation id0, amtfield4, id1, amtfield2
	task.AddRelation id1, amtfield2, id2, amtfield7
	task.IncludeAllFields
	task.CreateVirtualDatabase = False
	dbName = client.UniqueFilename(newFilename)
	Set fname5 = dbName
	task.OutputDatabaseName = dbName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set id0 = Nothing
	Set id1 = Nothing
	Set id2 = Nothing
	Client.opendatabase(dbName)
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
	Set db = Nothing
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
	Set db = Nothing
	Set field = Nothing
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

