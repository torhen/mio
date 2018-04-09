dim g_is_running
dim doc
dim app
dim fso

proj_path = "C:\proj" & "\"

open_atoll_project
refresh_and_overwrite
run_predictions "export_tp"
export_results "export_tp"
save_document
close_application

sub open_atoll_project
	' if the ATL was deleted, recover it 
	Set fso = CreateObject("Scripting.FileSystemObject")
	If not fso.FileExists(proj_path & "proj.atl") Then
		'Wscript.echo "Copy proj.doc to proj.atl"
		fso.CopyFile proj_path & "proj.doc", proj_path & "proj.atl", True
	End If

	' Start Atoll and load document
	set app = CreateObject("Atoll.Application")
	wscript.ConnectObject app, "Atoll_" ' Otherwise the events will not be cached!!!
	app.Visible = True
	set doc = app.Documents.Open(proj_path & "proj.atl")
	
end sub

sub refresh_and_overwrite
	doc.refresh 0   'cancel changes and reload database!!!
	
	' Overwrite some columns of Atoll tables
	overwrite_table  "ltransmitters", "C:\proj\trx\update_ltransmitters.csv"
	overwrite_table  "lcells",        "C:\proj\trx\update_lcells.csv"
	overwrite_table  "utransmitters", "C:\proj\trx\update_utransmitters.csv"
	overwrite_table  "gtransmitters", "C:\proj\trx\update_gtransmitters.csv"
	overwrite_table  "grepeaters",    "C:\proj\trx\update_grepeaters.csv"
	
	app.LogMessage "Overwrite completed."
end sub

sub run_predictions(pred_folder)
	unlock_pred pred_folder
	run_pred
end sub

sub export_results(pred_folder)
	export_result pred_folder, proj_path & "export"
end sub

sub save_document
	doc.Save()
end sub
	
sub close_application
	Wscript.DisconnectObject app
	doc = Null
	app.Documents.CloseAll 0
	app.Quit 0 
	app = Null

	' backup the ATL, because they are deleted regulary by a script
	fso.CopyFile proj_path & "proj.atl", proj_path & "proj.doc", True

end sub


private sub fill_dict(dict, file_name)

	set fso = CreateObject("Scripting.FileSystemObject")
	set file = fso.OpenTextFile(file_name)

	do until file.AtEndOfStream
		line = file.ReadLine
		if not dict.exists(line) then
			dict.add line , 1
		end if
	loop

end sub


private sub set_lte_load(load_value)
	app.LogMessage "Script set lte load " & load_value

	set tab = doc.GetRecords("lcells", True)
	i_all = tab.RowCount
	
	dim arr()
	redim arr(i_all,2)
	
	arr(0,0) = "dl_load"
	arr(0,1) = "ul_load"
	arr(0,2) = "TABULARDATA_POSITION"
	
	for i=1 to i_all
		arr(i,0) = load_value
		arr(i,1) = load_value
		arr(i,2) = i
	next

	res = tab.SetValues(arr)

end sub

private sub unlock_pred(folder_name)

	'first lock all predictions
	set root_folder =  doc.GetRootFolder(0).Item("Predictions")
	for each pred_folder in root_folder
		for each pred in pred_folder
			pred.SetProperty "LOCKED", True
		next
	next
	

	'now unlock predictions
	set pred_folder =  doc.GetRootFolder(0).Item("Predictions").Item(folder_name)
	
	for each pred in pred_folder
		app.LogMessage "Script unlocking " & pred.name
		pred.SetProperty "LOCKED", False
	next 

end sub

sub Atoll_RunComplete(arg1, arg2)
	g_is_running = 0
end sub


sub export_result(pred_folder, dest_folder)
	'prefix of prediction name defines export type

	set oCS = doc.CoordSystemDisplay

	set preds = doc.GetRootFolder(0).Item("Predictions").Item(pred_folder)

	for each pred in preds

		pred_name = pred.name
		
		if not fso.FolderExists(dest_folder) then
			MsgBox "Prediction destination folder " & dest_folder & "does not exist."
			exit sub
		end if

		app.LogMessage  "Script export " & pred_name
		
		file_type = lcase(right(pred_name,3))
		
		'text reported is always generated
		
		file_name = pred_name
		file_name = replace(file_name, " ", "_")
		file_name = replace(file_name, ":", "_")
		
		' make visible
		pred.Visible = True
		
		pred.export dest_folder & "\" & file_name & ".txt"  , oCS, "TXT"
		
		if file_type = "tab" then
			pred.export dest_folder & "\" & file_name, oCS, "TAB"
		end if
		
		if file_type = "mif" then
			pred.export dest_folder & "\" & file_name, oCS, "MIF"
		end if
		
		if file_type = "asc" then
			pred.export dest_folder & "\" & file_name, oCS, "ARCVIEWGRIDASCII"
		end if
		
	next
end sub

private sub run_pred
	app.LogMessage "Script run predictions"
	doc.run False
	g_is_running = 1
	do while g_is_running = 1
		wscript.sleep 1000
	loop
end sub

function csv2dict(csv_file)
	sep = ";"
	Set fso = CreateObject("Scripting.FileSystemObject")

	Set inputFile = fso.OpenTextFile(csv_file)

	Set dict = CreateObject("Scripting.Dictionary")
	
	i=0
	Do While inputFile.AtEndOfStream <> True
		i = i+1
		sLine = inputFile.ReadLine

		if i=1 then
			' save the header with a special key
			sLine = sLine + sep + "TABULARDATA_POSITION"
			aLine = split(sLine, sep)
			dict.add "_header_", aLine
		else
			' fill the dicionary
			sLine = sLine + sep + "-1"
			aLine = split(sLine, sep)
			dict.add aLine(0), aLine
		end if

	Loop
	
	Set csv2dict = dict
end function
	
sub overwrite_table(atoll_table, csv_file)

	app.LogMessage "overwrite " & atoll_table
	
	' Get data from csv
	Set dict = csv2dict(csv_file)
	
	' find name of primary key column
	sPrime = dict.item("_header_")(0)
	

	' Header
	aHeader = dict.item("_header_")
	iCols = ubound(aHeader)
	
	' get primary key data from atoll table
	set tab = doc.GetRecords(atoll_table, True)
	dim cols(0)
	cols(0) = aHeader(0)
	aPrimeData = tab.GetValues(empty,cols)
	iRows = tab.RowCount

	' create the input array
	dim input_arr()
	redim input_arr(iRows, iCols-1)	
	
	' set the header
	for i=1 to iCols:
		input_arr(0, i-1) = rtrim(aHeader(i))
	next
	
	' set the data
	for r = 1 to iRows
		PrimeKey = aPrimeData(r,1)
		
		if not dict.Exists(PrimeKey) then
			'app.LogMessage "not found " & PrimeKey
			PrimeKey = "_default_"
		end if
		
		for c=1 to iCols:
			input_arr(r, c-1) = rtrim(dict.item(PrimeKey)(c))
		next
		input_arr(r, iCols-1) = r


	next
		
	
	tab.SetValues(input_arr)
	
end sub







	