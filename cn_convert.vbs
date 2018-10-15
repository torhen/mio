option explicit

class ExcelWorkBook

	dim g_app, g_wb

	function init()
		Set g_app = CreateObject("Excel.Application")
		g_app.Visible = True
		g_app.DisplayAlerts = True
		g_app.WindowState = -4143
	end function

	function create_new()
		Set g_wb = g_app.Workbooks.Add()
	end function

	function open(file)
		Set g_wb = g_app.Workbooks.Open(file)
	end function

	function quit()
		g_app.quit
		set g_app = Nothing
	end function

	function get_value(sheet, range_str)
		dim s
		set s =  g_wb.Sheets(sheet)
		get_value = s.range(range_str).value
	end function

	function set_value(sheet, range_str, new_value)
		dim s
		set s =  g_wb.Sheets(sheet)
		s.range(range_str).value = new_value
	end function

	function sel(sheet, range_str)
		dim s
		set s =  g_wb.Sheets(sheet)
		s.activate
		s.range(range_str).select
	end function

	function import(file, sheet, after_sheet)
		dim w_tmp : set w_tmp = g_app.Workbooks.Open(file, False, True)
		dim s_tmp : Set s_tmp = w_tmp.Sheets(sheet)
		dim s : set s = g_wb.Sheets(after_sheet)
		s_tmp.copy(s)
		w_tmp.saved = True
		w_tmp.close
	end function

	function save()
		g_wb.save
	end function

	function save_as(file)
		g_app.DisplayAlerts = False
		g_wb.SaveAs file
		g_app.DisplayAlerts = True
	end function

	function delete_rows(sheet, range_str)
	' range_str in format 10:32;39:43;74:77 SEMIKOLONS!
		dim s
		set s =  g_wb.Sheets(sheet)
		s.Range(range_str).Delete
	end function

	function delete_sheet(sheet)
		g_wb.Sheets(sheet).Delete
	end function

	function replace_by_values(sheet)
		dim s
		set s =  g_wb.Sheets(sheet)
		s.activate
		s.Cells.copy
		s.Cells.pasteSpecial -4163 
		s.range("A1").select
	end function

	function find_rows(sheet, range_str, search_str, how)
		dim s, rg ,c, row, firstAddress, ub, res()

		if how = "whole" then how = 1 end if
		if how = "part"  then how = 2 end if

		set s =  g_wb.Sheets(sheet)
		set rg = s.Range(range_str)
		Set c = rg.Find(search_str, rg.cells(1,1), -4163,  how)

		ub = 0
		If Not c Is Nothing Then 
		    firstAddress = c.Address 
		    Do 
		    	ReDim Preserve res(ub)
		    	res(ub) = c.row
		    	ub = ub + 1
		    	
		        Set c = rg.FindNext(c) 
		    Loop While Not c Is Nothing And c.Address <> firstAddress 
		End If 

		find_rows = join(res, ";")

	end function


	function set_xy(sheet, rows, cols, set_value)
		dim s, str, col, row
		rows = split(rows, ";")
		cols = split(cols, ";")

		set s =  g_wb.Sheets(sheet)

		str = ""
		for each col in cols
			for each row in rows
				 str =  str + col + row + ";"
			Next
		next

		if  str <> "" then
			 str = left( str, len( str)-1)  ' strip the semicolon'
			if set_value= "" then
				s.Range(str).ClearContents
			else
				s.Range(str).value = set_value
			end if
		end if

	end function

	function basename(path)
		dim a : a = split(path, "\")
		basename = a(ubound(a))
	end function

end class


function create_excel(source_file, template_file, dest_file):

	' --- Update Template ----
	dim wb_tmpl
	set wb_tmpl = new ExcelWorkBook
	wb_tmpl.init()
	wb_tmpl.open(template_file)

	dim job
	job = wb_tmpl.basename(source_file)
	job = left(job, 10)
	wb_tmpl.set_value "Cover","E14", job
	wb_tmpl.set_value "Cover","U2", date

	dim group :group = wb_tmpl.get_value("Cover", "L14")

	wb_tmpl.save()
	wb_tmpl.quit()	

	if group <> "Swiss Towers - Sunrise" then
		log source_file &  " is not a Swiss Tower Site, no file created."
		exit function
	end if

	' --- Create Excel File ----
	dim wb 
	set wb = new ExcelWorkBook
	wb.init()
	wb.create_new()
	wb.import template_file,  "Cover", "Sheet1"
	wb.import source_file,  "Site solution", "Sheet1"
	wb.delete_sheet("Sheet1")
	wb.replace_by_values("Cover")
	wb.replace_by_values("Site solution")

	wb.delete_rows "Site solution", "10:32;40:44;75:78"

	dim rows_str
	rows_str = wb.find_rows("Site solution", "E:E", "RCK", "whole")
	wb.set_xy "Site solution", rows_str, "C;D;E" ,""
	rows_str = wb.find_rows("Site solution", "L:L", "RCK", "whole")
	wb.set_xy "Site solution", rows_str, "J;K;L" ,""
	rows_str = wb.find_rows("Site solution", "T:T", "RCK", "whole")
	wb.set_xy "Site solution", rows_str, "R;S;T" ,""
	rows_str = wb.find_rows("Site solution", "AA:AA", "RCK", "whole")
	wb.set_xy "Site solution", rows_str, "Y;Z;AA" ,""


	rows_str = wb.find_rows("Site solution", "C11:C15", "RRU", "part")
	wb.set_xy "Site solution", rows_str,  "C" ,"RRU"
	rows_str = wb.find_rows("Site solution", "J11:J15", "RRU", "part")
	wb.set_xy "Site solution", rows_str, "J" ,"RRU"
	rows_str = wb.find_rows("Site solution", "R11:R15", "RRU", "part")
	wb.set_xy "Site solution", rows_str,  "R" ,"RRU"
	rows_str = wb.find_rows("Site solution", "Y11:Y15", "RRU", "part")
	wb.set_xy "Site solution", rows_str, "Y" ,"RRU"

	wb.sel "Cover", "A1"
	wb.save_as(dest_file)
	wb.quit()	

end function

function basename(path)
	dim a : a = split(path, "\")
	basename = a(ubound(a))
end function

function log(msg)
	wscript.StdOut.WriteLine msg
	If Not g_logfile_obj Is Nothing Then
		g_logfile_obj.WriteLine msg
	end if

end function

function glob(folder, pattern)
	Dim oFSO, file, re, dict
	Set re = New RegExp
	re.pattern = pattern
	re.IgnoreCase = True
	Set dict = CreateObject("Scripting.Dictionary")

	Set oFSO=CreateObject("Scripting.FileSystemObject")
	for each file in	oFSO.GetFolder(folder).files
		if re.Test( file.name ) then
			'log file
			dict.Add file, file.name
		end if
	next
	Set oFSO = Nothing
	set re = Nothing
	set glob = dict
end function

function read_text_file(file)
	dim s, a
	dim fso  : Set fso = CreateObject("Scripting.FileSystemObject")
	dim fin : Set fin = fso.OpenTextFile(file, 1)
	s = fin.Readall
	fin.close

	' get rid of empty lines
	s = replace(s, vbCrLf, " ")
	s = trim(s)

	a = Split(s, " ")

	set fin = nothing
	set fso = nothing
	read_text_file = a
end function

function make_ir(job, source_dir, template_file, dest_dir, pattern)
	dim file, src_file, src_files, i, base, dest_file, search_pattern

	job = trim(job)
	if len(job)<>10 then
		log "'" + job & "' is not a job"
		exit function
	end if
		
	search_pattern = job + pattern
	set src_files = glob(source_dir + "\", search_pattern)

	' find latest file
	src_file = ""
	i = 0
	for each file in src_files
		i = i + 1
		if file > src_file then
			src_file = file
		end if
	next

	log "found " & i & " source files."

	'Process src_file
	if len(src_file) > 5 then
		dest_file = dest_dir + "\" + "IR_" + basename(src_file)
		log "try to create " & dest_file
		create_excel src_file, template_file, dest_file
	end if	

end Function

function get_wd()
	'dim fso  : Set fso = CreateObject("Scripting.FileSystemObject")
	'get_wd = fso.GetAbsolutePathName(".") + "\"
	'set fso = nothing

	get_wd = "\\serne\cn_convert\"
end function

function get_timestamp()
	dim s
	s = FormatDateTime(now)
	s = mid(s,7, 4) & "-" & mid(s, 4, 2) & "-" & mid(s, 1,2) & "_" & mid(s, 12, 2) & mid(s, 15, 2) & mid(s, 18, 2)
	get_timestamp = s

end function

function create_logfile()
	dim s, logWrite, log

	s = get_timestamp()
	s = get_wd() & "log\" & s & ".txt" 

    Set logWrite = CreateObject("Scripting.FileSystemObject")
    wscript.StdOut.WriteLine "Create Logfile: " & s
    set g_logfile_obj = logWrite.CreateTextFile(s, True)
    

end function

dim g_logfile_obj

function main(source_dir, dest_dir, jobs_txt, pattern)
	dim i, iall, job, jobs

	create_logfile()

	jobs = read_text_file(get_wd + jobs_txt)
	i=0
	iall = ubound(jobs) + 1
	for each job in jobs
		i = i + 1
		log "" & i & "/" & iall & " search for " & job
		make_ir job, source_dir, get_wd + "Template.xlsx", get_wd + dest_dir, pattern
	next
	log "finished " & get_timestamp()  

end function


main "\\swi.srse.net\dfs\Info\TE\Hua_Eng\RE\4.5G and TK Site solution\Detail Design", "dest", "jobs.txt", ".+v3.+\.xlsx"
