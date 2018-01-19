dim g_is_running
dim doc
dim app
dim fso

' start
make_reports


sub make_reports

	if wscript.arguments.count = 0 then
		MsgBox "Please supply commandline argument."
		exit sub
	end if

	proj_path = "C:\proj" & "\"
	
	' if the ATL was deleted, recover it 
	Set fso = CreateObject("Scripting.FileSystemObject")
	If not fso.FileExists(proj_path & "proj.doc") Then
		fso.CopyFile proj_path & "proj.doc",proj_path & "proj.atl", True
	End If

	' Start Atoll and load document
	set app = CreateObject("Atoll.Application")
	wscript.ConnectObject app, "Atoll_" ' Otherwise the evaents will not be cached!!!
	app.Visible = True
	set doc = app.Documents.Open(proj_path & "proj.atl")
	
	' to create cfg file:
	   ' - delete all macros 
	   ' - switch on all predictions (all visible)
	   ' - create one report
	   ' - save config file (all oprions selected) to report.cfg
	doc.SetConfig ( "C:\proj\data\report.cfg")
	
	' start the process dependend on commandline argument
	Set args = Wscript.Arguments
	arg = args(0)

	if arg = "set_values" then
		doc.refresh
		set dict = CreateObject("Scripting.Dictionary")
		fill_dict dict, proj_path & "trx_active.csv"
		set_trx dict, "gtransmitters"
		set_trx dict, "utransmitters"
		set_trx dict, "ltransmitters"
		set_lte_load 20
		set_gsm_rep 33
	else
		pred_folder = arg
		unlock_pred pred_folder
		run_pred
		export_result pred_folder, proj_path & pred_folder
	end if

	doc.Save()
	Wscript.DisconnectObject app
	doc = Null
	app.Documents.CloseAll 0
	app.Quit 0 
	app = Null

	' backup the ATL, because they are deleted reulary by a script
	fso.CopyFile proj_path & "proj.atl", proj_path & "proj.doc", True
	fso.CopyFile proj_path & "make_reports.vbs", proj_path & "make_reports.doc", True


	Wscript.Quit 0

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


private sub set_trx(dict, table_name)
	app.LogMessage "Script set trx " & table_name

	set tab = doc.GetRecords(table_name, True)
	i_all = tab.RowCount
	
	' get the transmittes names from table
	a = tab.GetValues(empty,empty) ' get the whole table
	dim arr_trx()
	redim arr_trx(i_all)
	for i = 1 to i_all
		arr_trx(i) = a(i,2)
	next

	' set values
	dim arr()
	redim arr(i_all,5)
	
	arr(0,0) = "active"
	arr(0,1) = "propag_model"
	arr(0,2) = "calc_radius"
	arr(0,3) = "calc_resolution"
	arr(0,4) = "propag_model2"
	arr(0,5) = "TABULARDATA_POSITION"
	
	for i=1 to i_all
		trx = arr_trx(i)
		arr(i,0) = dict.exists(trx)
		arr(i,1) = "CrossWave_SR"
		arr(i,2) = 25000
		arr(i,3) = 100
		arr(i,4) = ""
		arr(i,5) = i
	next

	res = tab.SetValues(arr)

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

private sub set_gsm_rep(eirp_value)
	app.LogMessage "Script set gsm rep " & eirp_value

	set tab = doc.GetRecords("grepeaters", True)
	i_all = tab.RowCount

	dim arr()
	redim arr(i_all,1)
	
	arr(0,0) = "eirp"
	arr(0,1) = "TABULARDATA_POSITION"
	
	for i=1 to i_all
		arr(i,0) = eirp_value
		arr(i,1) = i
	next

	res = tab.SetValues(arr)

end sub



private sub unlock_pred(folder_name)

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








	