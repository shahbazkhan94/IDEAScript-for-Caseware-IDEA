'****************************************************************************************************************
'* Script:		getFilename.iss
'* By:		Brain 
'* Version:	1.0
'* Date:		 
'* Purpose:	To get only file name of a file from the full path
'***************************************************************************************************************

'Test Case for getFileName function
Sub Main
	Dim path As String
	Set path = "C:\Users\Shahbaz Khan\Desktop\IDEA_Training_Data\Custom IDEAScript\JE_Working_final-v1.0.iss"	
	MsgBox  getFileName(path,0), MB_OK
End Sub

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
