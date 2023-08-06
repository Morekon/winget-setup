[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework")
[void] [Reflection.Assembly]::LoadWithPartialName("PresentationCore")

$Form = New-Object System.Windows.Forms.Form
$Form.Size = New-Object System.Drawing.Size(1200, 500)
$Form.StartPosition = "CenterScreen" #loads the window in the center of the screen
$Form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow #modifies the window border
$Form.Text = "winget setup" #window description
$Form.ShowInTaskbar = $True
$Form.KeyPreview = $True
$Form.AutoSize = $True
$Form.FormBorderStyle = 'Fixed3D'
$Form.MaximizeBox = $False
$Form.MinimizeBox = $False
$Form.ControlBox = $True
$Form.Icon = $Icon

# TODO automatically install winget if not installed
# TODO improve UI
# TODO improve error handling
############################################## Start functions
function WinGetInstaller($Remove)
{
    try
    {
        $winGet = gci "C:\Program Files\WindowsApps" -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.name -like "AppInstallerCLI.exe" -or $_.name -like "WinGet.exe" } | Select-Object -ExpandProperty fullname -ErrorAction SilentlyContinue
        # If there are multiple versions, select latest
        if ($winGet.count -gt 1)
        {
            $winGet = $winGet[-1]
        }
        $winGetLoc = [string]((get-item $winGet).Directory.FullName)
        $install = {
            Write-Host "Installing $id ..."
            & "$winGetLoc\winget.exe" install --id =$id -e --silent --accept-source-agreements --accept-package-agreements
        }
        $uninstall = {
            Write-Host "Removing $id ..."
            & "$winGetLoc\winget.exe" uninstall --id =$id -e --silent
        }

        $ids = @()

        # Browsers
        $ids += 'Brave.Brave' | Where-Object { $bravecb.Checked }
        $ids += 'Google.Chrome' | Where-Object { $chromecb.Checked }
        $ids += 'Mozilla.Firefox' | Where-Object { $firefoxcb.Checked }
        $ids += 'TorProject.TorBrowser' | Where-Object { $torcb.Checked }
        $ids += 'Microsoft.Edge' | Where-Object { $edgecb.Checked }
        $ids += 'Opera.Opera' | Where-Object { $operacb.Checked }
        $ids += 'VivaldiTechnologies.Vivaldi' | Where-Object { $vivaldicb.Checked }

        # Comunications
        $ids += 'Telegram.TelegramDesktop' | Where-Object { $telegramcb.Checked }
        $ids += 'WhatsApp.WhatsApp' | Where-Object { $whatsappcb.Checked }
        $ids += 'SlackTechnologies.Slack' | Where-Object { $slackcb.Checked }
        $ids += 'Discord.Discord' | Where-Object { $discordcb.Checked }
        $ids += 'Mattermost.MattermostDesktop' | Where-Object { $mattermostcb.Checked }
        $ids += 'TeamSpeakSystems.TeamSpeakClient' | Where-Object { $teamspeakcb.Checked }
        $ids += 'Zoom.Zoom' | Where-Object { $zoomcb.Checked }
        $ids += 'Microsoft.Skype' | Where-Object { $skypecb.Checked }
        $ids += 'Microsoft.Teams' | Where-Object { $teamscb.Checked }
        $ids += 'AnyDeskSoftwareGmbH.AnyDesk' | Where-Object { $anydeskcb.Checked }
        $ids += 'TeamViewer.TeamViewer' | Where-Object { $teamviewercb.Checked }

        # Development
        $ids += 'Git.Git' | Where-Object { $gitcb.Checked }
        $ids += 'Fork.Fork' | Where-Object { $forkcb.Checked }
        $ids += 'JetBrains.Toolbox' | Where-Object { $jetbrainstoolboxcb.Checked }
        $ids += 'Microsoft.VisualStudioCode' | Where-Object { $vscodecb.Checked }
        $ids += 'Docker.DockerDesktop' | Where-Object { $dockercb.Checked }
        $ids += 'PostgresSQL.pgAdmin' | Where-Object { $pgadmin.Checked }
        $ids += 'Insomnia.Insomnia' | Where-Object { $insomniacb.Checked }
        $ids += 'OpenJS.NodeJS.LTS' | Where-Object { $nodejscb.Checked }
        $ids += 'Notepad++.Notepad++' | Where-Object { $notepadcb.Checked }
        $ids += 'RProject.R' | Where-Object { $rcb.Checked }
        $ids += 'Posit.RStudio' | Where-Object { $rstudiocb.Checked }
        $ids += 'Microsoft.WindowsTerminal' | Where-Object { $windowsterminalcb.Checked }
        $ids += 'Canonical.Ubuntu.2204' | Where-Object { $ubuntu2204cb.Checked }
        $ids += 'Oracle.JDK.20' | Where-Object { $jdk20cb.Checked }
        $ids += 'Python.Python.3.9' | Where-Object { $python39cb.Checked }
        #        $ids += 'PuTTY.PuTTY' | Where-Object { $puttycb.Checked }
        #        $ids += 'WinSCP.WinSCP' | Where-Object { $winscpcb.Checked }
        #        $ids += 'Microsoft.Sysinternals.TCPView' | Where-Object { $tcpviewcb.Checked }
        #        $ids += 'mRemoteNG.mRemoteNG' | Where-Object { $mremotengcb.Checked }
        #        $ids += 'Famatech.AdvancedIPScanner' | Where-Object { $ipscancb.Checked }
        #        $ids += 'WiresharkFoundation.Wireshark' | Where-Object { $wiresharkcb.Checked }

        # Multimedia
        $ids += 'dotPDNLLC.paintdotnet' | Where-Object { $paintnetcb.Checked } #seems to be broken as the download fails
        $ids += 'GIMP.GIMP' | Where-Object { $gimpcb.Checked }
        $ids += 'Inkscape.Inkscape' | Where-Object { $inkscapcb.Checked }
        $ids += 'VideoLAN.VLC' | Where-Object { $vlc.Checked }
        $ids += 'Spotify.Spotify' | Where-Object { $spotifycb.Checked }
        $ids += 'Amazon.Kindle' | Where-Object { $kindlecb.Checked }
        $ids += 'OBSProject.OBSStudio' | Where-Object { $obscb.Checked }

        # Office
        $ids += 'Microsoft.OneNote' | Where-Object { $onenotecb.Checked }
        $ids += 'Microsoft.OneDrive' | Where-Object { $onedrivecb.Checked }
        $ids += 'Google.Drive' | Where-Object { $googledrivecb.Checked }
        $ids += 'Dropbox.Dropbox' | Where-Object { $dropboxcb.Checked }
        $ids += 'Nextcloud.NextcloudDesktop' | Where-Object { $nextcloudcb.Checked }
        $ids += 'ownCloud.ownCloudDesktop' | Where-Object { $owncloudcb.Checked }
        $ids += 'Obsidian.Obsidian' | Where-Object { $obsidiancb.Checked }
        $ids += 'DigitalScholar.Zotero' | Where-Object { $zoterocb.Checked }
        # Mendeley not available on winget
        $ids += 'Adobe.Acrobat.Reader.64-bit' | Where-Object { $adobereadercb.Checked }
        $ids += 'Microsoft.Office' | Where-Object { $microsoftofficecb.Checked }
        $ids += 'TheDocumentFoundation.LibreOffice' | Where-Object { $libreofficecb.Checked }
        $ids += 'appmakes.Typora' | Where-Object { $typoracb.Checked }
        $ids += 'Mozilla.Thunderbird' | Where-Object { $thunderbirdcb.Checked }
        $ids += 'Toggl.TogglTrack' | Where-Object { $togglcb.Checked } # Does not work
        $ids += 'MiKTeX.MiKTeX' | Where-Object { $miktexcb.Checked }
        $ids += 'MiKTeX.MiKTeX' | Where-Object { $pandocguicb.Checked } # PandocGui depends on it
        $ids += 'JohnMacFarlane.Pandoc' | Where-Object { $pandoccb.Checked }
        $ids += 'JohnMacFarlane.Pandoc' | Where-Object { $pandocguicb.Checked } # PandocGui depends on it
        $ids += 'Ombrelin.PandocGui' | Where-Object { $pandocguicb.Checked } # Depends on Pandoc and MiKTeX

        # Games
        $ids += 'Mojang.MinecraftLauncher' | Where-Object { $minecraftcb.Checked }
        $ids += 'Valve.Steam' | Where-Object { $steamcb.Checked }
        $ids += 'Peppy.Osu!' | Where-Object { $osucb.Checked }

        # Security
        $ids += 'KeePassXCTeam.KeePassXC' | Where-Object { $keepassxccb.Checked }
        $ids += 'SatoshiLabs.trezor-suite' | Where-Object { $trezorsuitecb.Checked }
        $ids += 'ExodusMovement.Exodus' | Where-Object { $exoduscb.Checked }
        $ids += 'Electrum.Electrum' | Where-Object { $electrumcb.Checked }
        $ids += 'Acronis.CyberProtectHomeOffice' | Where-Object { $acroniscb.Checked }
        $ids += 'Malwarebytes.Malwarebytes' | Where-Object { $malwarebytescb.Checked }

        # Utilities
        $ids += 'Microsoft.VCRedist.2015+.x64' | Where-Object { $wingetuicb.Checked } # Microsoft.VCRedist.2015+.x64 needs to be installed before to function correctly as some .dll file is missing
        $ids += 'SomePythonThings.WingetUIStore' | Where-Object { $wingetuicb.Checked }
        $ids += 'Chocolatey.Chocolatey' | Where-Object { $chocolateycb.Checked }
        $ids += 'Chocolatey.ChocolateyGUI' | Where-Object { $chocolateyguicb.Checked }
        $ids += 'Microsoft.PowerToys' | Where-Object { $powertoyscb.Checked }
        $ids += 'RARLab.WinRAR' | Where-Object { $winrarcb.Checked }
        $ids += 'AntibodySoftware.WizTree' | Where-Object { $wiztreecb.Checked }
        $ids += 'Piriform.Recuva' | Where-Object { $recuvacb.Checked }

        ForEach ($id in $ids | Select-Object -Unique)
        {
            if ($Remove)
            {
                Invoke-Command $uninstall
            }
            else
            {
                Invoke-Command $install
            }
        }
        Write-Host "Done."
    }
    #end try

    catch
    {
        $outputBox.text = "Operation could not be completed"
    }
}

function Create-GroupBox()
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [int]$X,
        [Parameter(Mandatory = $true)]
        [int]$Y,
        [Parameter(Mandatory = $true)]
        [int]$Width,
        [Parameter(Mandatory = $true)]
        [int]$Height,
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Form]$Form
    )
    $GroupBox = New-Object System.Windows.Forms.GroupBox
    $GroupBox.Location = New-Object System.Drawing.Size($X, $Y)
    $GroupBox.Size = New-Object System.Drawing.Size($Width, $Height)
    $GroupBox.Text = $Text
    $Form.Controls.Add($GroupBox)
    return $GroupBox
}

function Create-CheckBox()
{
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,
        [Parameter(Mandatory = $true)]
        [int]$X,
        [Parameter(Mandatory = $true)]
        [int]$Y,
        [Parameter(Mandatory = $true)]
        [int]$Width,
        [Parameter(Mandatory = $true)]
        [int]$Height,
        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.GroupBox]$Form
    )
    $CheckBox = New-Object System.Windows.Forms.CheckBox
    $CheckBox.Location = New-Object System.Drawing.Size($X, $Y)
    $CheckBox.Size = New-Object System.Drawing.Size($Width, $Height)
    $CheckBox.Text = $Text
    $Form.Controls.Add($CheckBox)
    return $CheckBox
}
############################################## end functions
############################################## Start group boxes
$Browsers = Create-GroupBox -Text "Browsers:" -X 10 -Y 10 -Width 130 -Height 250 -Form $Form
$Comunications = Create-GroupBox -Text "Comunications:" -X 150 -Y 10 -Width 140 -Height 250 -Form $Form
$Development = Create-GroupBox -Text "Development:" -X 300 -Y 10 -Width 140 -Height 250 -Form $Form
$Multimedia = Create-GroupBox -Text "Multimedia:" -X 450 -Y 10 -Width 140 -Height 250 -Form $Form
$Office = Create-GroupBox -Text "Office:" -X 600 -Y 10 -Width 140 -Height 250 -Form $Form
$Games = Create-GroupBox -Text "Games:" -X 750 -Y 10 -Width 140 -Height 250 -Form $Form
$Security = Create-GroupBox -Text "Security:" -X 900 -Y 10 -Width 140 -Height 250 -Form $Form
$Utilities = Create-GroupBox -Text "Utilities:" -X 1050 -Y 10 -Width 130 -Height 250 -Form $Form
############################################## end group boxes
############################################## Start Browsers checkboxes
$bravecb = Create-CheckBox -Text "Brave" -X 10 -Y 20 -Width 120 -Height 20 -Form $Browsers
$chromecb = Create-CheckBox -Text "Chrome" -X 10 -Y 40 -Width 120 -Height 20 -Form $Browsers
$firefoxcb = Create-CheckBox -Text "Firefox" -X 10 -Y 60 -Width 120 -Height 20 -Form $Browsers
$torcb = Create-CheckBox -Text "Tor" -X 10 -Y 80 -Width 120 -Height 20 -Form $Browsers
$egdecb = Create-CheckBox -Text "Edge" -X 10 -Y 100 -Width 120 -Height 20 -Form $Browsers
$operacb = Create-CheckBox -Text "Opera" -X 10 -Y 120 -Width 120 -Height 20 -Form $Browsers
$vivaldicb = Create-CheckBox -Text "Vivaldi" -X 10 -Y 140 -Width 120 -Height 20 -Form $Browsers
############################################## End Browser checkboxes
############################################## Start Communications checkboxes
$telegramcb = Create-CheckBox -Text "Telegram" -X 10 -Y 20 -Width 120 -Height 20 -Form $Comunications
$whatsappcb = Create-CheckBox -Text "WhatsApp" -X 10 -Y 40 -Width 120 -Height 20 -Form $Comunications
$slackcb = Create-CheckBox -Text "Slack" -X 10 -Y 60 -Width 120 -Height 20 -Form $Comunications
$discordcb = Create-CheckBox -Text "Discord" -X 10 -Y 80 -Width 120 -Height 20 -Form $Comunications
$mattermostcb = Create-CheckBox -Text "Mattermost" -X 10 -Y 100 -Width 120 -Height 20 -Form $Comunications
$teamspeakcb = Create-CheckBox -Text "TeamSpeak" -X 10 -Y 120 -Width 120 -Height 20 -Form $Comunications
$zoomcb = Create-CheckBox -Text "Zoom" -X 10 -Y 140 -Width 120 -Height 20 -Form $Comunications
$skypecb = Create-CheckBox -Text "Skype" -X 10 -Y 160 -Width 120 -Height 20 -Form $Comunications
$teamscb = Create-CheckBox -Text "Teams" -X 10 -Y 180 -Width 120 -Height 20 -Form $Comunications
$anydeskcb = Create-CheckBox -Text "AnyDesk" -X 10 -Y 200 -Width 120 -Height 20 -Form $Comunications
$teamviewercb = Create-CheckBox -Text "TeamViewer" -X 10 -Y 220 -Width 120 -Height 20 -Form $Comunications
############################################## End Communications checkboxes
############################################## Start Development checkboxes
$gitcb = Create-CheckBox -Text "Git" -X 10 -Y 20 -Width 120 -Height 20 -Form $Development
$forkcb = Create-CheckBox -Text "Fork" -X 10 -Y 40 -Width 120 -Height 20 -Form $Development
$jetbrainstoolboxcb = Create-CheckBox -Text "JetBrains Toolbox" -X 10 -Y 60 -Width 120 -Height 20 -Form $Development
$vscodecb = Create-CheckBox -Text "VS Code" -X 10 -Y 80 -Width 120 -Height 20 -Form $Development
$dockercb = Create-CheckBox -Text "Docker" -X 10 -Y 100 -Width 120 -Height 20 -Form $Development
$pgadmin = Create-CheckBox -Text "pgAdmin" -X 10 -Y 120 -Width 120 -Height 20 -Form $Development
$insomniacb = Create-CheckBox -Text "Insomnia" -X 10 -Y 140 -Width 120 -Height 20 -Form $Development
$nodejscb = Create-CheckBox -Text "NodeJS" -X 10 -Y 160 -Width 120 -Height 20 -Form $Development
$notepadcb = Create-CheckBox -Text "Notepad++" -X 10 -Y 180 -Width 120 -Height 20 -Form $Development
$rcb = Create-CheckBox -Text "R" -X 10 -Y 200 -Width 120 -Height 20 -Form $Development
$rstudiocb = Create-CheckBox -Text "RStudio" -X 10 -Y 220 -Width 120 -Height 20 -Form $Development
$windowsterminalcb = Create-CheckBox -Text "Windows Terminal" -X 10 -Y 240 -Width 120 -Height 20 -Form $Development
$ubuntu2204cb = Create-CheckBox -Text "Ubuntu 20.04" -X 10 -Y 260 -Width 120 -Height 20 -Form $Development
$jdk20cb = Create-CheckBox -Text "JDK 20" -X 10 -Y 280 -Width 120 -Height 20 -Form $Development
$python39cb = Create-CheckBox -Text "Python 3.9" -X 10 -Y 300 -Width 120 -Height 20 -Form $Development
############################################## End Development  checkboxes
############################################## Start Multimedia checkboxes
$paintnetcb = Create-CheckBox -Text "Paint.NET" -X 10 -Y 20 -Width 120 -Height 20 -Form $Multimedia
$gimpcb = Create-CheckBox -Text "GIMP" -X 10 -Y 40 -Width 120 -Height 20 -Form $Multimedia
$inkscapecb = Create-CheckBox -Text "Inkscape" -X 10 -Y 60 -Width 120 -Height 20 -Form $Multimedia
$vlccb = Create-CheckBox -Text "VLC" -X 10 -Y 80 -Width 120 -Height 20 -Form $Multimedia
$spotifycb = Create-CheckBox -Text "Spotify" -X 10 -Y 100 -Width 120 -Height 20 -Form $Multimedia
$kindlecb = Create-CheckBox -Text "Kindle" -X 10 -Y 120 -Width 120 -Height 20 -Form $Multimedia
$obscb = Create-CheckBox -Text "OBS" -X 10 -Y 140 -Width 120 -Height 20 -Form $Multimedia
############################################## End Multimedia checkboxes
############################################## Start Office checkboxes
$onedrivecb = Create-CheckBox -Text "OneDrive" -X 10 -Y 20 -Width 120 -Height 20 -Form $Office
$googledrivecb = Create-CheckBox -Text "Google Drive" -X 10 -Y 40 -Width 120 -Height 20 -Form $Office
$dropboxcb = Create-CheckBox -Text "Dropbox" -X 10 -Y 60 -Width 120 -Height 20 -Form $Office
$nextcloudcb = Create-CheckBox -Text "Nextcloud" -X 10 -Y 80 -Width 120 -Height 20 -Form $Office
$owncloudcb = Create-CheckBox -Text "ownCloud" -X 10 -Y 100 -Width 120 -Height 20 -Form $Office
$obsidiancb = Create-CheckBox -Text "Obsidian" -X 10 -Y 120 -Width 120 -Height 20 -Form $Office
$zoterocb = Create-CheckBox -Text "Zotero" -X 10 -Y 140 -Width 120 -Height 20 -Form $Office
$adobereadercb = Create-CheckBox -Text "Adobe Reader" -X 10 -Y 160 -Width 120 -Height 20 -Form $Office
$microsoftofficecb = Create-CheckBox -Text "Microsoft Office" -X 10 -Y 180 -Width 120 -Height 20 -Form $Office
$libreofficecb = Create-CheckBox -Text "LibreOffice" -X 10 -Y 200 -Width 120 -Height 20 -Form $Office
$typoracb = Create-CheckBox -Text "Typora" -X 10 -Y 220 -Width 120 -Height 20 -Form $Office
$thunderbirdcb = Create-CheckBox -Text "Thunderbird" -X 10 -Y 240 -Width 120 -Height 20 -Form $Office
$togglcb = Create-CheckBox -Text "Toggl" -X 10 -Y 260 -Width 120 -Height 20 -Form $Office
$miktexcb = Create-CheckBox -Text "MiKTeX" -X 10 -Y 280 -Width 120 -Height 20 -Form $Office
$pandoccb = Create-CheckBox -Text "Pandoc" -X 10 -Y 300 -Width 120 -Height 20 -Form $Office
$pandocguicb = Create-CheckBox -Text "PandocGui" -X 10 -Y 320 -Width 120 -Height 20 -Form $Office
############################################## End Office checkboxes
############################################## Start Games checkboxes
$minecraftcb = Create-CheckBox -Text "Minecraft" -X 10 -Y 20 -Width 120 -Height 20 -Form $Games
$steamcb = Create-CheckBox -Text "Steam" -X 10 -Y 40 -Width 120 -Height 20 -Form $Games
$osucb = Create-CheckBox -Text "osu!" -X 10 -Y 60 -Width 120 -Height 20 -Form $Games
############################################## End Games checkboxes
############################################## Start Security checkboxes
$keepassxcb = Create-CheckBox -Text "KeePassXC" -X 10 -Y 20 -Width 120 -Height 20 -Form $Security
$trezorsuitecb = Create-CheckBox -Text "Trezor Suite" -X 10 -Y 40 -Width 120 -Height 20 -Form $Security
$exoduscb = Create-CheckBox -Text "Exodus" -X 10 -Y 60 -Width 120 -Height 20 -Form $Security
$electrumcb = Create-CheckBox -Text "Electrum" -X 10 -Y 80 -Width 120 -Height 20 -Form $Security
$acroniscb = Create-CheckBox -Text "Acronis" -X 10 -Y 100 -Width 120 -Height 20 -Form $Security
$malwarebytescb = Create-CheckBox -Text "Malwarebytes" -X 10 -Y 120 -Width 120 -Height 20 -Form $Security
############################################## End Security checkboxes
############################################## Start Ultilities checkboxes
$wingetuicb = Create-CheckBox -Text "WingetUI" -X 10 -Y 20 -Width 120 -Height 20 -Form $Utilities
$chocolateycb = Create-CheckBox -Text "Chocolatey" -X 10 -Y 40 -Width 120 -Height 20 -Form $Utilities
$chocolateyguicb = Create-CheckBox -Text "Chocolatey GUI" -X 10 -Y 60 -Width 120 -Height 20 -Form $Utilities
$powertoyscb = Create-CheckBox -Text "PowerToys" -X 10 -Y 80 -Width 120 -Height 20 -Form $Utilities
$winrarcb = Create-CheckBox -Text "WinRAR" -X 10 -Y 100 -Width 120 -Height 20 -Form $Utilities
$wiztreecb = Create-CheckBox -Text "WizTree" -X 10 -Y 120 -Width 120 -Height 20 -Form $Utilities
$recuvacb = Create-CheckBox -Text "Recuva" -X 10 -Y 140 -Width 120 -Height 20 -Form $Utilities
############################################## End Ultilities checkboxes
############################################## Start buttons
$submitInstallButton = New-Object System.Windows.Forms.Button
$submitInstallButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$submitInstallButton.BackColor = [System.Drawing.Color]::LightGreen
$submitInstallButton.Location = New-Object System.Drawing.Size(10, 280)
$submitInstallButton.Size = New-Object System.Drawing.Size(110, 40)
$submitInstallButton.Text = "Install "
$submitInstallButton.Add_Click({ WinGetInstaller })
$Form.Controls.Add($submitInstallButton)

$submitUninstallButton = New-Object System.Windows.Forms.Button
$submitUninstallButton.Cursor = [System.Windows.Forms.Cursors]::Hand
$submitUninstallButton.BackColor = [System.Drawing.Color]::Red
$submitUninstallButton.Location = New-Object System.Drawing.Size(140, 280)
$submitUninstallButton.Size = New-Object System.Drawing.Size(110, 40)
$submitUninstallButton.Text = "Uninstall "
$submitUninstallButton.Add_Click({ WinGetInstaller $true })
$Form.Controls.Add($submitUninstallButton)
############################################## end buttons
$Form.Add_Shown({ $Form.Activate() })
[void] $Form.ShowDialog()
