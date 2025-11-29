$recentFolder = [Environment]::GetFolderPath("Recent")

Write-Host ""
Write-Host "@@@@@@    @@@@@@      @@@       @@@@@@@@   @@@@@@   @@@@@@@   @@@  @@@     @@@  @@@@@@@  "
Write-Host "@@@@@@@   @@@@@@@      @@@       @@@@@@@@  @@@@@@@@  @@@@@@@@  @@@@ @@@     @@@  @@@@@@@  "
Write-Host "!@@       !@@          @@!       @@!       @@!  @@@  @@!  @@@  @@!@!@@@     @@!    @@!    "
Write-Host "!@!       !@!          !@!       !@!       !@!  @!@  !@!  @!@  !@!!@!@!     !@!    !@!    "
Write-Host "!!@@!!    !!@@!!       @!!       @!!!:!    @!@!@!@!  @!@!!@!   @!@ !!@!     !!@    @!!    "
Write-Host " !!@!!!    !!@!!!      !!!       !!!!!:    !!!@!!!!  !!@!@!    !@!  !!!     !!!    !!!    "
Write-Host "     !:!       !:!     !!:       !!:       !!:  !!!  !!: :!!   !!:  !!!     !!:    !!:    "
Write-Host "    !:!       !:!       :!:      :!:       :!:  !:!  :!:  !:!  :!:  !:!     :!:    :!:    "
Write-Host ":::: ::   :::: ::       :: ::::   :: ::::  ::   :::  ::   :::   ::   ::      ::     ::  "
Write-Host ":: : :    :: : :       : :: : :  : :: ::    :   : :   :   : :  ::    :      :       :"
Write-Host ""
Write-Host "https://discord.gg/UET6TdxFUk"
Write-Host ""

$totalWidth = 60
$headerText = "SHELL:RECENT VIEWER"
$padding = [math]::Floor(($totalWidth - $headerText.Length) / 2)
Write-Host ("-" + ("=" * ($totalWidth - 2)) + "-") -ForegroundColor DarkGray
Write-Host ("|" + (" " * $padding) + $headerText + (" " * ($totalWidth - 2 - $padding - $headerText.Length)) + "|") -ForegroundColor DarkGray
Write-Host ("-" + ("=" * ($totalWidth - 2)) + "-") -ForegroundColor DarkGray
Write-Host ""

Write-Host "[LIST] Elenco dei file aperti di recente tramite Shell:Recent" -ForegroundColor Cyan
Write-Host "Percorso cartella: $recentFolder" -ForegroundColor Yellow
Write-Host "Ora attuale: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Yellow
Write-Host ("-" * $totalWidth)
Write-Host ""

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
                Write-Host "[FOUND] File esistente" -ForegroundColor Green
                Write-Host "  ? Percorso : $targetPath" -ForegroundColor Yellow
                Write-Host "  ? Modifica : $writeTime" -ForegroundColor DarkCyan
                Write-Host "  ? Accesso  : $accessTime" -ForegroundColor DarkCyan
            } else {
                Write-Host "[NOT FOUND] File NON trovato" -ForegroundColor Red
                Write-Host "  ? Percorso : $targetPath" -ForegroundColor DarkRed
                Write-Host "  ? Modifica : $writeTime" -ForegroundColor DarkCyan
                Write-Host "  ? Accesso  : $accessTime" -ForegroundColor DarkCyan
            }
            Write-Host ("-" * $totalWidth)
        }
    }

Write-Host ""
$footerText = "OPERAZIONE COMPLETATA"
$padding = [math]::Floor(($totalWidth - $footerText.Length) / 2)
Write-Host ("-" + ("=" * ($totalWidth - 2)) + "-") -ForegroundColor DarkGray
Write-Host ("|" + (" " * $padding) + $footerText + (" " * ($totalWidth - 2 - $padding - $footerText.Length)) + "|") -ForegroundColor DarkGray
Write-Host ("-" + ("=" * ($totalWidth - 2)) + "-") -ForegroundColor DarkGray
Write-Host ""

