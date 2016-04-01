Sub Main
	Call AppendDatabase()	'test3-Sheet1.IMD
End Sub


' File: Append Databases
Function AppendDatabase
	Set db = Client.OpenDatabase("test1-Sheet1.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "test2-Sheet1.IMD"
	task.AddDatabase "test3-Sheet1.IMD"
	dbName = "Append Databases.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function