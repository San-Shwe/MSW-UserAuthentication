dim app, wks, db, strBE, strFE

' strFE = CurrentProject.Path & "\" & process_pvc.accdb 'to use in project app

strDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strFE = strDir & "\UserAuthorization_be.accdb"
'strBE = strDir & "\UserAuthorization_be.accdb"

strPVCBE =  "C:\Kaizaki\Document\MS Access\MSW-PROCESS-PVC\process_pvc_be.accdb"
strOPPBE =  "C:\Kaizaki\Document\MS Access\MSW-PROCESS-OPP\process_opp_be.accdb"
strUserAuthBE = "C:\Kaizaki\Document\MS Access\MSW-UserAuthentication\UserAuthorization_be.accdb"

set app = CreateObject("Access.Application")
set wks = app.dbEngine.createworkspace("", "admin", "", 2) 'dbUseJet
set db = wks.opendatabase(strFE)

LinkMyTable db, strOPPBE, "FG", "FG"
LinkMyTable db, strOPPBE, "OF", "OPP_OF"
LinkMyTable db, strOPPBE, "tblOPPSummary", "tblOPPSummary"

LinkMyTable db, strPVCBE, "FG", "FG"
LinkMyTable db, strPVCBE, "OF", "PVC_OF"
LinkMyTable db, strPVCBE, "tblPVCSummary", "tblPVCSummary"

LinkMyTable db, strUserAuthBE, "tblEmployeeAccess", "tblEmployeeAccess"
LinkMyTable db, strUserAuthBE, "tblEmployees", "tblEmployees"
LinkMyTable db, strUserAuthBE, "tblUser", "tblUser"

db.close
app.quit
msgbox "Your Application was updated", vbInformation, "Updated"
'cleanup 
set db=Nothing
set wks = Nothing
set app = Nothing

Sub DeleteOldTable(db, strTable)
	'Delete an old table 
	on Error Resume Next
	db.tabledefs.delete strTable
end sub

sub LinkMyTable(db, strBE, strSource, strAlias)
	'Link a table to a backend file 
	dim strConnect 
	dim tdf 
	DeleteOldTable db, strAlias
	strConnect	 = ";DATABASE=" & strBE
	Set tdf = db.createtabledef(strAlias, 0, strSource, strConnect)
	db.tabledefs.Append tdf
	set tdf = Nothing
	
End sub
