'****************************************************************************************************************
'* Script:		summarization_func-generic.iss
'* By:		Shahbaz Khan
'* Version:	1.0
'* Date:		March 24, 2016
'* Purpose:	for using summarization task in idea
'* Usability:	Can only be used with other complete tests having dialogs. 
'***************************************************************************************************************
Sub Main
	Dim file As String
	
	Set file = " " 'enter file name from the current working directory
	Call Summarization() 'Use with any file
	Client.refereshfileexplorer
End Sub

' Analysis: Summarization
Function Summarization
	Dim db As database
	Dim task As task
	Dim dbName As String
	
	Set db = Client.OpenDatabase(filename) 'Filename editable
	Set task = db.Summarization
	task.UseQuickSummarization = TRUE 	'Change for using more than one fields for summarization
	task.AddFieldToSummarize "ACCOUNTNO"		 'Fieldname used for summarization
	task.AddFieldToTotal "ACCOUNTED_DR"  		' Filedname for total on, can add more fields
	task.AddFieldToTotal "ACCOUNTED_CR"
	dbName = "Summarized JE by Account.IMD"	'New db name'
	task.OutputDBName = dbName			
	task.CreatePercentField = FALSE			'used to show percetage with summarization
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	
	'Clearing memory
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
