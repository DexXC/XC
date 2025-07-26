Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Sub Log vide pour ignorer les logs
Sub Log(msg)
    ' Pas de log
End Sub

On Error Resume Next

Log "=== Démarrage du script de désactivation du reset Windows ==="

' Désactivation de WinRE
Log "Désactivation de WinRE..."
objShell.Run "cmd /c reagentc /disable", 0, True

' Suppression des dossiers Recovery
If fso.FolderExists("C:\Recovery") Then
    Log "Suppression de C:\Recovery..."
    objShell.Run "cmd /c rd /s /q C:\Recovery", 0, True
End If

If fso.FolderExists("C:\$SysReset") Then
    Log "Suppression de C:\$SysReset..."
    objShell.Run "cmd /c rd /s /q C:\$SysReset", 0, True
End If

' Registre : désactivation du reset
Log "Désactivation via registre..."
objShell.Run "cmd /c reg add ""HKLM\Software\Policies\Microsoft\Windows\System"" /v ""DisableResetToDefault"" /t REG_DWORD /d 1 /f", 0, True

' BCD : désactivation du recovery
Log "Désactivation via BCD..."
objShell.Run "cmd /c bcdedit /set {current} recoveryenabled No", 0, True

' Fonction pour supprimer les partitions Recovery
Sub RemoverParticoesRecovery(diskNumber)
    Log "Analyse du disque " & diskNumber

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
        If (InStr(UCase(line), "RECUP") > 0) _
        Or (InStr(UCase(line), "RECOVERY") > 0) _
        Or (InStr(line, "恢复") > 0) Then
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

' Balayage des disques 0 à 9 (inclus Thaïlande, Chine, etc.)
For d = 0 To 9
    RemoverParticoesRecovery d
Next

Log "=== Fin du script ==="
