Dim strfolder
Dim lngwarning : lngwarning = 8000
Dim lngcritic : lngcritic = 9000
Dim wsh
Dim lngvelkost
Dim lngjednotka
Dim Perf_Data
Dim dbNames
Dim i : i = 0

'##########################################################'
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set wsh = CreateObject("WScript.Shell")
'##########################################################'

const MANDATORY_DB = "<your primary database>"
const HKEY_LOCAL_MACHINE = &H80000002
const REG_INSTANCE_NAMES = "Software\Microsoft\Microsoft SQL Server\Instance Names\SQL"
const REG_INSTANCE_NAME_VALUE = "SQLEXPRESS"
const REG_SQL_PATH = "Software\Microsoft\Microsoft SQL Server\"
const REG_SQL_DATA_VALUE = "SQLDataRoot"
const DATA_DIR = "DATA"
const MDF_EXTENSION = "mdf"
const LOGFILE_ENDING = "_log.ldf"
strComputer = "." 

Set regObj=GetObject( _ 
    "winmgmts:{impersonationLevel=impersonate}!\\" & _
   strComputer & "\root\default:StdRegProv")

' extract the instance name from registry
regObj.GetStringValue HKEY_LOCAL_MACHINE,REG_INSTANCE_NAMES,REG_INSTANCE_NAME_VALUE,strValue

' extract the file system path where the data files reside
strKeyPath = REG_SQL_PATH & strValue & "\Setup" 
regObj.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,REG_SQL_DATA_VALUE,strValue

strFolder = strValue & "\" & DATA_DIR

' warning and critical values may get overwritten by arguments
If Wscript.Arguments.Count = 2 Then
  lngwarning = Wscript.Arguments(0)
  lngcritic = Wscript.Arguments(1)
End if

Recurse objFSO.GetFolder(strFolder)

Sub Recurse(objFolder)
    Dim objFile

    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = MDF_EXTENSION Then
            totalSize = objFile.Size
            logFile = strFolder & "\" & objFSO.GetBaseName(objFile) & LOGFILE_ENDING
            If (objFSO.FileExists(logFile)) Then
                ' system databases use a different naming scheme, but are very small so can be ignored
                logSize = objFSO.GetFile(logFile).Size
                totalSize = totalSize + logSize
            End If
            if (totalSize/1024000) > CLng(lngcritic) then 
                Wscript.Echo "CRITICAL: " & round (totalSize / 1048576,1) & " MB " & objFSO.GetBaseName(objFile)
                Wscript.Quit(2)
            elseif (totalSize/1048576) > CLng(lngwarning) then 
                Wscript.Echo "WARNING: " & round (totalSize / 1048576,1) & " MB " & objFSO.GetBaseName(objFile)
                Wscript.Quit(1)
            else
                if i=0 then
                    dbNames = objFSO.GetBaseName(objFile)
                else
                    dbNames = dbNames & ", " & objFSO.GetBaseName(objFile)
                End If
            End If
        End If
        i = i + 1
    Next
End Sub

If InStr(dbNames, MANDATORY_DB) = 0 Then
  Wscript.Echo "CRITICAL: Database " & MANDATORY_DB & " not found"
  Wscript.Quit(2)
End If

Wscript.Echo "OK: " & dbNames
Wscript.Quit(0)
