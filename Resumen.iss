Sub Main
	Call Summarization()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Análisis: Resumen
Function Summarization
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "NUM_VENDEDOR"
	task.AddFieldToSummarize "COD_PROD"
	task.AddFieldToSummarize "NUM_CLI"
	task.AddFieldToTotal "NETO"
	task.AddFieldToTotal "IMPUESTO"
	task.AddFieldToTotal "TOTAL"
	dbName = "Resumen01.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function