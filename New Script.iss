Begin Dialog newdlg 66,20,284,241,"Bulk Excel Import", .newdlg
  PushButton 224,12,40,14, "Browse", .PushButton1
  Text 16,14,195,21, "Select the path", .Text1
  OKButton 166,201,40,14, "OK", .OKButton1
  CancelButton 224,201,40,14, "Cancel", .CancelButton1
  TextBox 15,57,242,119, .TextBox1
  TextBox 17,199,140,15, .TextBox2
  Text 18,188,40,8, "New Filename", .Text2
  Text 18,45,40,9, "Processing", .Text2
End Dialog
Option Explicit


Sub Main

End Sub


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

'getFilename function

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

