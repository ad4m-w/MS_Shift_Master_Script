# Created By Adam Waszczyszak
# Scripts Disabled Bypass from CMD: powershell -ExecutionPolicy Bypass -File "C:\Temp\MSShell.ps1"

# Self-check for admin rights, and ask for perms if launched not as admin (from Superuser.com)
param([switch]$Elevated)

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}

# Clear Screen after admin check
cls

# Menu from StackOverflow
Function MenuMaker{
    param(
        [parameter(Mandatory=$true)][String[]]$Selections,
        [switch]$IncludeExit,
        [string]$Title = $null
        )

    $Width = if($Title){$Length = $Title.Length;$Length2 = $Selections|%{$_.length}|Sort -Descending|Select -First 1;$Length2,$Length|Sort -Descending|Select -First 1}else{$Selections|%{$_.length}|Sort -Descending|Select -First 1}
    $Buffer = if(($Width*1.5) -gt 78){[math]::floor((78-$width)/2)}else{[math]::floor($width/4)}
    if($Buffer -gt 6){$Buffer = 6}
    $MaxWidth = $Buffer*2+$Width+$($Selections.count).length+2
    $Menu = @()
    $Menu += "╔"+"═"*$maxwidth+"╗"
    if($Title){
        $Menu += "║"+" "*[Math]::Floor(($maxwidth-$title.Length)/2)+$Title+" "*[Math]::Ceiling(($maxwidth-$title.Length)/2)+"║"
        $Menu += "╟"+"─"*$maxwidth+"╢"
    }
    For($i=1;$i -le $Selections.count;$i++){
        $Item = "$(if ($Selections.count -gt 9 -and $i -lt 10){" "})$i`. "
        $Menu += "║"+" "*$Buffer+$Item+$Selections[$i-1]+" "*($MaxWidth-$Buffer-$Item.Length-$Selections[$i-1].Length)+"║"
    }
    If($IncludeExit){
        $Menu += "║"+" "*$MaxWidth+"║"
        $Menu += "║"+" "*$Buffer+"X - Exit"+" "*($MaxWidth-$Buffer-8)+"║"
    }
    $Menu += "╚"+"═"*$maxwidth+"╝"
    $menu
}

MenuMaker -Selections 'Create Temp and Visitor_pics Folders',
'Set Temp and Visitor_Pics Permissions', 
'Set PTI Folder Permissions',
'Block DYMO Updates', 
'Block Adobe Auto-Update Service', 
'Stop Print Spooler', 
'Restart Print Spooler', 
'Delete all print jobs on PC', 
'Install Drivers via elevated path', 
'Download and Install API', 
'Download and Install Adobe', 
'Download and Install Signature Pad', 
'Download and Install HF Scanner (DS8101) + PDFs',
'Download and Install HF Scanner (DS6707) + PDFs',
'Download and Install LX 500 driver',
'Download and Install DYMO 550 driver',
'Download and Install DYMO 450 driver',
'Download and Install GK420d driver',
'Download and Install GC420d driver (download broken, use GK420d)',
'Download and Install ZD421/420 driver',
'Download and Install ZXP-7 driver',
'Change API Port', 
'Launch Print Server', 
'Microsoft Edge Registry Patch (Edge Engine Error)', 
'Create Download Links and set Clip Board' -Title 'Choose a Function' -IncludeExit

$MenuChoice = Read-Host "Choose an option"

while($MenuChoice -ne 'X'){

    if($MenuChoice -eq 'X'){
        Write-Output "Exiting Menu..."
        break
        exit
    }

    if($MenuChoice -eq 1){
           if (Test-Path -Path C:\Temp){
            "Temp Folder Already Exists"
        }

            else{
                New-Item -Path C:\Temp -ItemType Directory
            }

            if (Test-Path -Path C:\Visitor_pic){
                "Visitor_Pic Folder Already Exists"
            }

            else{
                New-Item -Path C:\Visitor_pic -ItemType Directory
            }
        }

        if($MenuChoice -eq 2){         
            
        $path=Get-Acl -Path C:\Temp
        $acl=New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule ('BUILTIN\Users','FullControl','ContainerInherit, ObjectInherit','None','Allow')
        $path.setaccessrule($acl)
        Set-Acl -Path C:\Temp\ -AclObject $path


        $path=Get-Acl -Path C:\Visitor_pic
        $acl=New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule ('BUILTIN\Users','FullControl','ContainerInherit, ObjectInherit','None','Allow')
        $path.setaccessrule($acl)
        Set-Acl -Path C:\Visitor_pic\ -AclObject $path

        "Permissions for Temp and Visitor_Pics Set!"
    }
        
        if($MenuChoice -eq 3){ 

            $path=Get-Acl -Path C:\ProgramData\PTI
            $acl=New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule ('BUILTIN\Users','FullControl','ContainerInherit, ObjectInherit','None','Allow')
            $path.setaccessrule($acl)
            Set-Acl -Path C:\ProgramData\PTI\ -AclObject $path

            "PTI Folder Permissions Set!"
    }

#    if($MenuChoice -eq 4){
#        Write-Output "Sending Test Print"
#            $printerName = "Color Label 500" 
#            Get-CimInstance Win32_Printer -Filter "name LIKE '%$printerName%'" |
#                        Invoke-CimMethod -MethodName printtestpage
#    }
#
#        if($MenuChoice -eq 5){
#            Write-Output "Sending Test Print"
#              $printerName = "DYMO LabelWriter 550 Turbo" 
#             Get-CimInstance Win32_Printer -Filter "name LIKE '%$printerName%'" |
#                       Invoke-CimMethod -MethodName printtestpage
#    }
#        if($MenuChoice -eq 6){
#            Write-Output "Sending Test Print"
#               $printerName = "Zebra ZXP Series 7 USB Card Printer" 
#                Get-CimInstance Win32_Printer -Filter "name LIKE '%$printerName%'" |
#                        Invoke-CimMethod -MethodName printtestpage
#    }
        if($MenuChoice -eq 4){

            New-NetFirewallRule -Program "C:\Program Files (x86)\DYMO\DYMO Connect\DYMOConnect.exe" -Action Block -Profile Domain, Private, Public -DisplayName “Block DYMO Connect” -Description “Block DYMO Connect” -Direction Outbound

            New-NetFirewallRule -Program "C:\Program Files (x86)\DYMO\DYMO Connect\DYMO.WebApi.Win.Host.exe" -Action Block -Profile Domain, Private, Public -DisplayName “Block DYMO WebService” -Description “Block DYMO WebService” -Direction Outbound
            "Services Blocked In Firewall"
        }
        if($MenuChoice -eq 5){
        
        Set-Service -Name "AdobeARMservice" -StartupType Disabled

        "Adobe Update Services Blocked In Services.msc"
    }
        if($MenuChoice -eq 6){
        
            Stop-Service -Name Spooler
            'Print Spooler Service Stopped'
        }

        if($MenuChoice -eq 7){

            Start-Service -Name Spooler
            'Print Spooler Service Started'
        }

        if($MenuChoice -eq 8){

            Get-Printer | ForEach-Object { Remove-PrintJob -PrinterName $_.Name }
            'All print jobs deleted'
        }

        if($MenuChoice -eq 9){
            
            $local = Read-Host "Path to .exe"
            Start-Process -FilePath “$local”
            "Success!"

        }

        if($MenuChoice -eq 10){
            
            'Parsing download site...'
            
            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/e2d06108-8cc8-4705-a316-54463dc1d78f"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText

            'Downloading...'
                   
            $Destination = "C:\Temp\api.zip" 
            Invoke-WebRequest -Uri $text -OutFile $Destination
            'Uncompressing...'
            Expand-Archive -LiteralPath 'C:\Temp\api.zip' -DestinationPath C:\Temp
            "Launching API installer..."
            Start-Process -FilePath "C:\Temp\MSShift.DevicesAPI.Setup.NEW.msi"
            "Success!"

        }

        if($MenuChoice -eq 11){
            
            'Parsing download site...'

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/5da99203-21ba-4aa2-93e6-a60a8a0b3ae3"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading...'

            $Destination = "C:\Temp\adobe.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching Adobe installer..."
            Start-Process -FilePath "C:\Temp\adobe.exe"
            "Success!"

        }

        if($MenuChoice -eq 12){
            
            'Parsing download site...'

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/e43d957b-0b20-4422-a3d0-a114162c5dfe"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading...'

            $Destination = "C:\Temp\sigplus_.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching Signature Pad installer..."
            Start-Process -FilePath "C:\Temp\sigplus_.exe"
            "Success!"

        }

        if($MenuChoice -eq 13){
            
            'Parsing download site...'
            # Download HF driver
            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/30ce8450-e037-4228-8987-432968180d43"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading driver...'

            $Destination = "C:\Temp\Zebra_CoreScanner_Driver.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            #Downloading PDF's
            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/76cdef97-2774-4b11-9adb-14b0220159f5"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading Restore Default PDF...'

            $Destination = "C:\Temp\Restore Default.pdf" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/e5a5d996-6a66-400b-a2c7-548627642815"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading ScanX_Config_Codebar PDF...'

            $Destination = "C:\Temp\ScanX_Config_Codebar.pdf" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching Hands Free Scanner installer..."
            Start-Process -FilePath "C:\Temp\Zebra_CoreScanner_Driver.exe"
            "Success!"

        }

        if($MenuChoice -eq 14){
            
            'Parsing download site...'
            # Download HF driver
            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/be6e546f-2bff-4547-ad52-a13442f9a53f"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading driver...'

            $Destination = "C:\Temp\Zebra123_CoreScanner_Driver.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            #Downloading PDF's
            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/76cdef97-2774-4b11-9adb-14b0220159f5"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading Restore Default PDF...'

            $Destination = "C:\Temp\Restore Default.pdf" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/e5a5d996-6a66-400b-a2c7-548627642815"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading ScanX_Config_Codebar PDF...'

            $Destination = "C:\Temp\ScanX_Config_Codebar.pdf" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching Hands Free Scanner installer..."
            Start-Process -FilePath "C:\Temp\Zebra123_CoreScanner_Driver.exe"
            "Success!"

        }

        if($MenuChoice -eq 15){
            
            'Parsing download site...'

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/7ebdd547-4c3c-4dc4-8639-e0ce88c1f60c"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading...'

            $Destination = "C:\Temp\Primera.2.3.1.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching LX 500 installer..."
            Start-Process -FilePath "C:\Temp\Primera.2.3.1.exe"
            "Success!"

        }
        if($MenuChoice -eq 16){
            
            'Parsing download site...'

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/1ab37806-4228-4eb1-8178-1ba492b0ea0f"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading...'

            $Destination = "C:\Temp\DCDSetup1.4.5.1.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching DYMO 550 driver installer..."
            Start-Process -FilePath "C:\Temp\DCDSetup1.4.5.1.exe"
            "Success!"

        }

        if($MenuChoice -eq 17){
            
            'Parsing download site...'

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/2f89909d-4539-446b-a76c-1ff7f47954aa"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading...'

            $Destination = "C:\Temp\DCDSetup1.3.2.18.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching DYMO 450 driver installer..."
            Start-Process -FilePath "C:\Temp\DCDSetup1.3.2.18.exe"
            "Success!"

        }

        if($MenuChoice -eq 18){
            
            'Parsing download site...'

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/2923c379-b7a0-4506-9c28-4ea5b2c0e48c"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading...'

            $Destination = "C:\Temp\zd51.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching GK420d driver installer..."
            Start-Process -FilePath "C:\Temp\zd51.exe"
            "Success!"

        }

        if($MenuChoice -eq 19){
            
            'Parsing download site...'

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/56bbaadf-adb3-4a2b-875b-68dac3bb2489"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading...'

            $Destination = "C:\Temp\zd.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching GC420d driver installer..."
            Start-Process -FilePath "C:\Temp\zd.exe"
            "Success!"

        }

        if($MenuChoice -eq 20){
            
            'Parsing download site...'

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/75b8f783-51f0-4a8a-9128-bf6957480aa4"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading...'

            $Destination = "C:\Temp\zd105.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching ZD421/420 driver installer..."
            Start-Process -FilePath "C:\Temp\zd105.exe"
            "Success!"

        }

        if($MenuChoice -eq 21){
            
            'Parsing download site...'

            # Retrieve the HTML content of the website
            $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/7cc7f865-5107-4f8b-9f3e-617b8ca23802"
            # Extract the text content from the parsed HTML
            $text = $response.ParsedHtml.body.innerText
            
            'Downloading...'

            $Destination = "C:\Temp\ZXP73.0.2.exe" 
            Invoke-WebRequest -Uri $text -OutFile $Destination

            "Launching ZXP-7 driver installer..."
            Start-Process -FilePath "C:\Temp\ZXP73.0.2.exe"
            "Success!"

        }

        if($MenuChoice -eq 22){
            
            "Manually input the old port number"
            $content = Get-Content "C:\Program Files (x86)\MS Shift Inc\MS Shift DevicesAPI\MSShift.DevicesAPI.exe.config"
            'Current API Port #: '
            (Get-Content -Path "C:\Program Files (x86)\MS Shift Inc\MS Shift DevicesAPI\MSShift.DevicesAPI.exe.config" -TotalCount 55)[-1]
            $oldPort = Read-Host "Old Port Number"            
            $newPort = Read-Host "New Port Number"
            $content = $content -replace "$oldPort", "$newPort"
            Set-Content "C:\Program Files (x86)\MS Shift Inc\MS Shift DevicesAPI\MSShift.DevicesAPI.exe.config" -Value $content
            "Changed API port to: $newPort"

        }
        if($MenuChoice -eq 23){
            'Launching print server'
            rundll32 printui.dll, PrintUIEntry /s 

        }

        if($MenuChoice -eq 24){

        $regContent = @"

        Windows Registry Editor Version 5.00
        ;Disable IE11 Welcome Screen
        [HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Main]
        "DisableFirstRunCustomize"=dword:00000001 
"@

        # Create a temporary .reg file
        $tempFile = New-TemporaryFile
        $tempFile.FullName
        $regContent | Out-File -FilePath $tempFile.FullName -Encoding ASCII

        # Import the registry settings
        reg.exe import $tempFile.FullName

        # Clean up the temporary file
        Remove-Item $tempFile.FullName           

        }

        if($MenuChoice -eq 25){

            '1.API'
            '2.Adobe'
            '3.Signature Pad'
            '4.HF Scanner (DS8101)'
            '5.HF Scanner (DS6707)'
            '6.LX 500'
            '7.DYMO 550'
            '8.DYMO 450'
            '9.GK420d'
            '10.ZD421/420'
            '11.ZXP-7'

            $driverChoice = Read-Host "Please Choose a Driver"  
         
            if($driverChoice -eq 1){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/e2d06108-8cc8-4705-a316-54463dc1d78f"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }
            if($driverChoice -eq 2){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/5da99203-21ba-4aa2-93e6-a60a8a0b3ae3"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }
            if($driverChoice -eq 3){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/e43d957b-0b20-4422-a3d0-a114162c5dfe"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }

            if($driverChoice -eq 4){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/30ce8450-e037-4228-8987-432968180d43"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }

            if($driverChoice -eq 5){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/30ce8450-e037-4228-8987-432968180d43"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }

            if($driverChoice -eq 6){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/7ebdd547-4c3c-4dc4-8639-e0ce88c1f60c"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }
            if($driverChoice -eq 7){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/1ab37806-4228-4eb1-8178-1ba492b0ea0f"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }
            if($driverChoice -eq 8){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/2f89909d-4539-446b-a76c-1ff7f47954aa"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }

            if($driverChoice -eq 9){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/2923c379-b7a0-4506-9c28-4ea5b2c0e48c"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }

            if($driverChoice -eq 10){

                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/75b8f783-51f0-4a8a-9128-bf6957480aa4"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }
            if($driverChoice -eq 11){
                
                # Retrieve the HTML content of the website
                $response = Invoke-WebRequest -Uri "https://download.msshift.com/link/7cc7f865-5107-4f8b-9f3e-617b8ca23802"
                # Extract the text content from the parsed HTML
                $text = $response.ParsedHtml.body.innerText
                Set-Clipboard -Value "$text"
                'Driver download link copied to clipboard'
                'Exiting Driver Selection...'
            }
        }
        
    $MenuChoice = Read-Host "Choose another menu option"

}

Write-Output "Exiting Powershell..."