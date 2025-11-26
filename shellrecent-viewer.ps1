$recentFolder = [Environment]::GetFolderPath("Recent")

Write-Output "@@@@@@    @@@@@@      @@@       @@@@@@@@   @@@@@@   @@@@@@@   @@@  @@@     @@@  @@@@@@@  "
Write-Output "@@@@@@@   @@@@@@@      @@@       @@@@@@@@  @@@@@@@@  @@@@@@@@  @@@@ @@@     @@@  @@@@@@@  "
Write-Output "!@@       !@@          @@!       @@!       @@!  @@@  @@!  @@@  @@!@!@@@     @@!    @@!    "
Write-Output "!@!       !@!          !@!       !@!       !@!  @!@  !@!  @!@  !@!!@!@!     !@!    !@!    "
Write-Output "!!@@!!    !!@@!!       @!!       @!!!:!    @!@!@!@!  @!@!!@!   @!@ !!@!     !!@    @!!    "
Write-Output " !!@!!!    !!@!!!      !!!       !!!!!:    !!!@!!!!  !!@!@!    !@!  !!!     !!!    !!!    "
Write-Output "     !:!       !:!     !!:       !!:       !!:  !!!  !!: :!!   !!:  !!!     !!:    !!:    "
Write-Output "    !:!       !:!       :!:      :!:       :!:  !:!  :!:  !:!  :!:  !:!     :!:    :!:    "
Write-Output ":::: ::   :::: ::       :: ::::   :: ::::  ::   :::  ::   :::   ::   ::      ::     ::  "  
Write-Output ":: : :    :: : :       : :: : :  : :: ::    :   : :   :   : :  ::    :      :       :"    
Write-Output ""
Write-Output "https://discord.gg/UET6TdxFUk"
Write-Output "" 

Write-Host "`n------------------------------------------" -ForegroundColor DarkGray
Write-Host "            SHELL:RECENT VIEWER"
Write-Host "------------------------------------------`n" -ForegroundColor DarkGray
Write-Output "" 
Write-Host "[LIST] Elenco dei file aperti di recente tramite Shell:Recent"
Write-Host ""
Write-Host " Percorso: $recentFolder"
Write-Host " Ora attuale:" $(Get-Date).ToString('dd/MM/yyyy HH:mm:ss')
Write-Host ""
Write-Host "------------------------------------------------------------------------"

Get-ChildItem -Path $recentFolder -Filter *.lnk -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    ForEach-Object {
        $lnkPath = $_.FullName
        $accessTime = $_.LastAccessTime.ToString('dd/MM/yyyy HH:mm:ss')
        $writeTime  = $_.LastWriteTime.ToString('dd/MM/yyyy HH:mm:ss')

        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut($lnkPath)
        $targetPath = $shortcut.TargetPath

        if (![string]::IsNullOrEmpty($targetPath)) {
            if (Test-Path $targetPath) {
                Write-Host " [FOUND] File esistente:" -ForegroundColor Green
                Write-Host " $targetPath" -ForegroundColor Yellow
                Write-Host "     Modifica: $writeTime"
                Write-Host "     Accesso : $accessTime"
            } else {
                Write-Host "[NOT FOUND] File NON trovato:" -ForegroundColor Red
                Write-Host " $targetPath" -ForegroundColor DarkRed
                Write-Host "     Ultima modifica: $writeTime"
                Write-Host "     Ultimo accesso: $accessTime"
            }
            Write-Host "------------------------------------------------------------------------"
        }
    }

Write-Host ""
Write-Host "[COMPLETE] Operazione completata con successo." -ForegroundColor Green
Write-Host ""







