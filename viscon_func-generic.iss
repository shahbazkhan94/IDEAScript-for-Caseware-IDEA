'****************************************************************************************************************
'* Script:		viscon_func-generic.iss
'* By:		Shahbaz Khan
'* Version:	1.0
'* Date:		March 24, 2016
'* Purpose:	for joining multiple files
'* Usability:	Can only be used with other complete tests having dialogs. 
'***************************************************************************************************************
' File: Visual Connector
Function RelateDatabase
	Dim db As database
	Dim task As task
	Dim dbName As String
	Dim id0 As Variant
	Dim id1 As Variant
	Dim id2 As Variant
	
	Set db = Client.OpenDatabase(fname2)		' primary filename
	Set task = db.VisualConnector
	id0 = task.AddDatabase (fname2)			' primary filename
	id1 = task.AddDatabase (fname1)			' secondary file 1
	id2 = task.AddDatabase (fname3)			' secondary file 2
	task.MasterDatabase = id0			
	task.AppendDatabaseNames = FALSE		' using database names
	task.IncludeAllPrimaryRecords = TRUE		
	task.AddRelation id0, "ACCOUNT_CODE", id1, "ACCOUNT_CODE"		'adding relations from primary to sec1 to sec2
	task.AddRelation id1, "ACCOUNT_CODE", id2, "ACCOUNTNO"
	task.IncludeAllFields				'include all fields
	task.CreateVirtualDatabase = False		
	dbName = "VisCon.IMD"				'new db name
	task.OutputDatabaseName = dbName		
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
