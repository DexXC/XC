Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

logPath = "C:\reset_desativacao.log"
Set logFile = fso.CreateTextFile(logPath, True)

Sub Log(msg)
    logFile.WriteLine Now & " - " & msg
End Sub

On Error Resume Next

Log "=== In�cio do script de desativa��o do reset do Windows ==="

' Desativar o WinRE
Log "Desativando o WinRE..."
objShell.Run "cmd /c reagentc /disable", 0, True

' Remover pastas de recupera��o
If fso.FolderExists("C:\Recovery") Then
    Log "Removendo C:\Recovery..."
    objShell.Run "cmd /c rd /s /q C:\Recovery", 0, True
End If

If fso.FolderExists("C:\$SysReset") Then
    Log "Removendo C:\$SysReset..."
    objShell.Run "cmd /c rd /s /q C:\$SysReset", 0, True
End If

' Registro: desativar redefini��o
Log "Desativando o reset via registro..."
objShell.Run "cmd /c reg add ""HKLM\Software\Policies\Microsoft\Windows\System"" /v ""DisableResetToDefault"" /t REG_DWORD /d 1 /f", 0, True

' BCD: desativar recupera��o
Log "Desativando o recovery via BCD..."
objShell.Run "cmd /c bcdedit /set {current} recoveryenabled No", 0, True

' Fun��o para remover parti��es de recupera��o
Sub RemoverParticoesRecovery(diskNumber)
    Log "Analisando disco " & diskNumber & " para parti��es de recupera��o..."

    tempScript = fso.GetSpecialFolder(2) & "\diskpart_" & diskNumber & ".txt"
    tempOut = fso.GetSpecialFolder(2) & "\diskpart_out_" & diskNumber & ".txt"

    ' Gerar script diskpart
    Set scriptFile = fso.CreateTextFile(tempScript, True)
    scriptFile.WriteLine "select disk " & diskNumber
    scriptFile.WriteLine "list partition"
    scriptFile.Close

    ' Executar diskpart
    objShell.Run "cmd /c diskpart /s """ & tempScript & """ > """ & tempOut & """", 0, True
    WScript.Sleep 2000

    If Not fso.FileExists(tempOut) Then
        Log "Nenhum resultado para disco " & diskNumber
        Exit Sub
    End If

    Set ts = fso.OpenTextFile(tempOut, 1)
    Do Until ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If (InStr(UCase(line), "RECUP") > 0) Or (InStr(UCase(line), "RECOVERY") > 0) Then
            partInfo = Split(line)
            If UBound(partInfo) >= 1 Then
                partNum = partInfo(1)
                Log "Parti��o de recupera��o detectada: disco " & diskNumber & ", parti��o " & partNum
                ' Gerar script de exclus�o
                Set delScript = fso.CreateTextFile(tempScript, True)
                delScript.WriteLine "select disk " & diskNumber
                delScript.WriteLine "select partition " & partNum
                delScript.WriteLine "delete partition override"
                delScript.Close
                objShell.Run "cmd /c diskpart /s """ & tempScript & """", 0, True
                WScript.Sleep 1000
                Log "Parti��o removida."
            End If
        End If
    Loop
    ts.Close

    fso.DeleteFile(tempScript)
    fso.DeleteFile(tempOut)
End Sub

' Executar para os discos 0 a 3
For d = 0 To 3
    RemoverParticoesRecovery d
Next

Log "=== Fim do script ==="
logFile.Close
