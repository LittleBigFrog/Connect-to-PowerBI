## Please provide Connection String
$ConnectionString =  "Provider=MSOLAP.8;Persist Security Info=True;Initial Catalog=sobe_wowvirtualserver-ee30a2fc-aeeb-4eb8-81d2-d1ec21eec060;Data Source=pbiazure://api.powerbi.com;MDX Compatibility=1;Safety Options=2;MDX Missing Member Mode=Error;Identity Provider=https://login.microsoftonline.com/common, https://analysis.windows.net/powerbi/api, 929d0ec0-7a41-4b1e-bc7c-b754a28bddcc;Update Isolation Level=2"

## Please provide SQL Query
$SQLquery = "Evaluate CandidatesData"
##############################################################################
$conn = New-Object System.Data.OleDb.OleDbConnection
$conn.ConnectionString = $ConnectionString
$comm = New-Object System.Data.OleDb.OleDbCommand($SQLquery,$conn)
$conn.Open()
Write-Output $comm
$adapter = New-Object System.Data.OleDb.OleDbDataAdapter $comm
$dataset = New-Object System.Data.DataSet
$adapter.Fill($dataSet)
$conn.Close()
$rows=($dataset.Tables | Select-Object -Expand Rows)
# Write-Output $rows

$table = $dataset.Tables[0]
$table | Export-CSV "out.csv"