' Copie este arquivo para a Area de Trabalho (ou qualquer pasta).
' Na primeira vez, abra o Fluxo uma vez a partir da pasta de instalacao.
' Depois disso, este atalho funciona de qualquer lugar.
'
' O Fluxo abre como qualquer planilha do Excel: nao fecha, move nem
' esconde as outras planilhas que voce ja tenha abertas.

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

Function NormalizeFolder(path)
    If path = "" Then
        NormalizeFolder = ""
        Exit Function
    End If
    NormalizeFolder = fso.GetAbsolutePathName(path)
End Function

Function IsValidInstall(folder)
    IsValidInstall = False
    If folder = "" Then Exit Function
    If fso.FileExists(folder & "\Fluxo_Caixa.xlsx") And fso.FileExists(folder & "\abrir-fluxo.ps1") Then
        IsValidInstall = True
    End If
End Function

Function ReadRegistryInstall()
    On Error Resume Next
    path = shell.RegRead("HKCU\Software\FluxoCaixa\InstallPath")
    On Error GoTo 0
    ReadRegistryInstall = NormalizeFolder(path)
End Function

Function ReadSidecarInstall(vbsFolder)
    ini = vbsFolder & "\fluxo-caixa.path"
    If Not fso.FileExists(ini) Then
        ReadSidecarInstall = ""
        Exit Function
    End If
    Set ts = fso.OpenTextFile(ini, 1, False)
    line = Trim(ts.ReadLine())
    ts.Close
    ReadSidecarInstall = NormalizeFolder(line)
End Function

vbsFolder = fso.GetParentFolderName(WScript.ScriptFullName)
installFolder = ""

If IsValidInstall(vbsFolder) Then
    installFolder = vbsFolder
Else
    sidecar = ReadSidecarInstall(vbsFolder)
    If IsValidInstall(sidecar) Then
        installFolder = sidecar
    Else
        reg = ReadRegistryInstall()
        If IsValidInstall(reg) Then
            installFolder = reg
        End If
    End If
End If

If installFolder = "" Then
    MsgBox "Nao foi possivel localizar o Fluxo de Caixa." & vbCrLf & vbCrLf & _
           "Abra uma vez a partir da pasta de instalacao" & vbCrLf & _
           "ou coloque um arquivo fluxo-caixa.path ao lado deste atalho" & vbCrLf & _
           "com o caminho da pasta (ex.: C:\FluxoCaixa).", vbCritical, "Fluxo de Caixa"
    WScript.Quit 1
End If

ps1 = installFolder & "\abrir-fluxo.ps1"
cmd = "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File """ & ps1 & """ -Silent -InstallPath """ & installFolder & """"
shell.Run cmd, 0, False
