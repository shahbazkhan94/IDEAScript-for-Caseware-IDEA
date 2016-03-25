'****************************************************************************************************************
'* Script:		join_func-generic.iss
'* Default:	Yes
'* By:		Shahbaz Khan
'* Purpose:	join two databases
'* Usability:	Can only be used with other  scripts
'***************************************************************************************************************

' File: Join Databases
Function JoinDatabase
	Dim db As database
	Dim task As task
	Dim dbName As String
	
	Set db = Client.OpenDatabase(fname2) 		'primary filename
	Set task = db.JoinDatabase			
	task.FileToJoin "Summarized JE by Account.IMD"	'secondary filename
	task.AddPFieldToInc "ACCOUNT_CODE"		'primary file name field to include
	task.AddPFieldToInc "DESCRIPTION"		'primary file name field to include
	task.AddPFieldToInc "CLOSING_DR"		'primary file name field to include
	task.AddPFieldToInc "CLOSING_CR"		'primary file name field to include
	task.AddSFieldToInc "ACCOUNTED_DR_SUM"	'secondary file name field to include
	task.AddSFieldToInc "ACCOUNTED_CR_SUM"	'secondary file name field to include
	task.AddMatchKey "ACCOUNT_CODE", "ACCOUNTNO", "A"  		'match key primary, secondary and order of precedence'
	task.CreateVirtualDatabase = False		
	dbName = "Join1.IMD"				'new db name
	task.PerformTask dbName, "", WI_JOIN_ALL_REC	'join type
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function