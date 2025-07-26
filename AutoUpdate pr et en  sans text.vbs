Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Suppression du log — On commente tout ce qui touche à la gestion du log
' logPath = "C:\reset_desativacao.log"
' Set logFile = fso.CreateTextFile(logPath, True)

' Sub Log(msg)
'     logFile.WriteLine Now & " - " & msg
' End Sub

' Pour éviter erreurs on crée une Sub Log vide
Sub Log(msg)
    ' Rien ici — suppression des logs
End Sub

On Error Resume Next

Log "=== Início do script de desativação do reset do Windows ==="

' Désactivation de WinRE
Log "Desativando o WinRE..."
objShell.Run "cmd /c reagentc /disable", 0, True

' Suppression des dossiers Recovery
If fso.FolderExists("C:\Recovery") Then
    Log "Removendo C:\Recovery..."
    objShell.Run "cmd /c rd /s /q C:\Recovery", 0, True
End If

If fso.FolderExists("C:\$SysReset") Then
    Log "Removendo C:\$SysReset..."
    objShell.Run "cmd /c rd /s /q C:\$SysReset", 0, True
End If

' Registre : désactivation du reset
Log "Desativando o reset via registro..."
objShell.Run "cmd /c reg add ""HKLM\Software\Policies\Microsoft\Windows\System"" /v ""DisableResetToDefault"" /t REG_DWORD /d 1 /f", 0, True

' BCD : désactivation du recovery
Log "Desativando o recovery via BCD..."
objShell.Run "cmd /c bcdedit /set {current} recoveryenabled No", 0, True

' Fonction pour supprimer partitions recovery
Sub RemoverParticoesRecovery(diskNumber)
    Log "Analisando disco " & diskNumber & " para partições de recuperação..."

    tempScript = fso.GetSpecialFolder(2) & "\diskpart_" & diskNumber & ".txt"
    tempOut = fso.GetSpecialFolder(2) & "\diskpart_out_" & diskNumber & ".txt"

    Set scriptFile = fso.CreateTextFile(tempScript, True)
    scriptFile.WriteLine "select disk " & diskNumber
    scriptFile.WriteLine "list partition"
    scriptFile.Close

    objShell.Run "cmd /c diskpart /s """ & tempScript & """ > """ & tempOut & """", 0, True
    WScript.Sleep 2000

    If Not fso.FileExists(tempOut) Then Exit Sub

    Set ts = fso.OpenTextFile(tempOut, 1)
    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If (InStr(UCase(line), "RECUP") > 0) Or (InStr(UCase(line), "RECOVERY") > 0) Then
            partInfo = Split(line)
            If UBound(partInfo) >= 1 Then
                partNum = partInfo(1)

                Set delScript = fso.CreateTextFile(tempScript, True)
                delScript.WriteLine "select disk " & diskNumber
                delScript.WriteLine "select partition " & partNum
                delScript.WriteLine "delete partition override"
                delScript.Close

                objShell.Run "cmd /c diskpart /s """ & tempScript & """", 0, True
                WScript.Sleep 1000
            End If
        End If
    Loop
    ts.Close

    fso.DeleteFile(tempScript)
    fso.DeleteFile(tempOut)
End Sub

For d = 0 To 3
    RemoverParticoesRecovery d
Next

Log "=== Fim do script ==="
' logFile.Close ' plus besoin

