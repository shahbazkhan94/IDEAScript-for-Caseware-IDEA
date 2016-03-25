'****************************************************************************************************************
'* Script:		menu_func.iss
'* By:		Shahbaz Khan
'* Version:	1.0
'* Date:		March 24, 2016
'* Purpose:	To ensure completeness of JEs by trial balance movement for millions of records
'* Usability:	Can only be used with other complete tests having dialogs. 
'***************************************************************************************************************

Function menu()
	'Variable definition for this function
	
	Dim dlg As dlgMenu
	Dim button As Integer
	Dim filebar As Object
	Dim exitDialog As Boolean
	Dim source As Object
	Dim table As Object
	Dim fields As Integer
	Dim i, j As Integer
	Dim field As Object
	
	'Starting loop for displaying dialog
	Do
	
		button = Dialog(dlg)
		
		'Using select .. case statement for different states of button
		
		Select Case button
			Case -1 'ok button 
				If dlg.DropListBox1 > -1 Then
					amtfield1 = ListBox1$(dlg.DropListBox1)
				Else
					amtfield1 = ""
				End If
											
				newFilename = dlg.txtNewFilename
				exitDialog = TRUE
			Case 0 ' cancel button
				exitDialog = TRUE
				exitScript = TRUE
			Case 1'filename select button
				
				'Using file explorer object
				Set filebar = CreateObject("ideaex.fileexplorer")
				filebar.displaydialog
				
				'File select within in file explorer
				fname1 = filebar.selectedfile
				
				'For getting table definition and getting numeric fields only
				If fname1 <> "" Then
					Set source = client.opendatabase(fname1)
					Set table = source.tabledef
					fields = table.count
					ReDim ListBox1$(fields)
					j = 0
					For i = 1 To fields
						Set field = table.getfieldat(i)
						If field.isnumeric Then ' rule to get numeric fields from table definition
							ListBox1$(j) = field.name
							j = j + 1
						End If
					Next i
					
				End If
				
				If j = 0 Then
					MsgBox "The file selected does not contain a numeric field", MB_ICONEXCLAMATION, "Error"
					fname1 = ""
				End If								
		End Select
	Loop While exitDialog = FALSE

	'Clearing memory
	
	Set source = Nothing
	Set table = Nothing
	Set field = Nothing
End Function
