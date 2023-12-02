# Device Settings - Checks audio and mouse configuration in Windows and applies your desired settings.
# V1.0.0 - 02/12/2023

# The AudioDeviceCmdlets module is required for checking and configuring audio device settings. https://github.com/frgnca/AudioDeviceCmdlets  
# To install this module the script must be run as Administrator for the first time it is launched.

###########################################################################
### Enter the desired devices/settings:

# Output Device (Headphones\Speakers)
$AudioOutput = "Headphones (2- AT2020USB+)"
$AudioOutputVolume = "80"

# Input Device (Microphone)
$Microphone = "Microphone (2- AT2020USB+)"
$MicrophoneVolume = "60"

# Mouse Settings
$MouseAcceleration = "0" # 0 = OFF, 1 = ON
$MousePointerSpeed = "10" # 10 = 6 in the Mouse Properties\Pointer Options tab.

###########################################################################




Function AudioDeviceCmdlets {

    ### Check AudioDeviceCmdlets Module is installed
    $Module = "AudioDeviceCmdlets"
    Import-Module -Name $Module
    IF (Get-Module -Name $Module) {
        
        # Get installed version number
        $InstalledVersion = Get-Module -Name $Module | Select-Object -ExpandProperty Version
        # Get latest version number
        $LatestVersion = Find-Module -Name $Module | Select-Object -ExpandProperty Version

        # Check Module is up to date and prompt to install if newer version is found.
        IF ([System.Version]$InstalledVersion -lt [System.Version]$LatestVersion) {

            $Response = Read-Host "A new version of $Module is available. Would you like to install? Y/N"
            IF ($Response -eq "Y") {

                try {

                    Update-Module -Name $Module

                }
                catch {

                    Write-Host $Error[0] -ForeGroundColor Red
                    Write-Output "Press any key to end the script..."
                    $Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown") | Out-Null
                    Exit

                }

            }

        }
    } ELSE {
        
        Write-Host "Module $Module not found. Attempting to install..."
        try {

            Install-Module -Name $Module

        } catch {
            
            Write-Host $Error[0] -ForeGroundColor Red
            Write-Output "Press any key to end the script..."
            $Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown") | Out-Null
            Exit

        }
        
    }

}




Function AudioOutputSettings {

    ### Check Audio Output Settings
    Write-Host "`nChecking Audio Output Settings..."
    # Check if the desired device is set as the default audio device
    $PlaybackDevice = Get-AudioDevice -list | Where-Object {$_.Name -eq $AudioOutput} | Select-Object Index, Name, Default, DefaultCommunication
    $PlaybackDevice | Select-Object Name, Default, DefaultCommunication | Format-List
    IF ($PlaybackDevice.Default -match "False") {

        $Response = $(Write-Host $AudioOutput "is not the default playback device. Set this as default? Y/N: " -ForegroundColor Yellow -NoNewLine; Read-Host)
        IF ($Response -eq "Y") {

            Set-AudioDevice $PlaybackDevice.Index -DefaultOnly | Out-Null
            Get-AudioDevice -index $PlaybackDevice.Index | Select-Object Name, Default, DefaultCommunication | Format-List

        }

    }
    # Check if the desired device is set as default playback communication device
    IF ($PlaybackDevice.DefaultCommunication -match "False") {

        $Response = $(Write-Host $AudioOutput "is not the default playback communication device. Set this as default? Y/N: " -ForegroundColor Yellow -NoNewLine; Read-Host)
        IF ($Response -eq "Y") {

            Set-AudioDevice $PlaybackDevice.Index -CommunicationOnly | Out-Null
            Get-AudioDevice -index $PlaybackDevice.Index | Select-Object Name, Default, DefaultCommunication | Format-List

        } 

    }
    # Check playback volume
    $PlaybackVolume = Get-AudioDevice -PlaybackVolume
    IF ($PlaybackVolume.Trim("%") -eq $AudioOutputVolume) {

        Write-Host "PlaybackVolume ="$PlaybackVolume -ForeGroundColor Green

    } ELSE {

        $Response = $(Write-Host "PlaybackVolume ="$PlaybackVolume"`nSet PlaybackVolume to" $AudioOutputVolume"%? Y/N: " -ForegroundColor Yellow -NoNewLine; Read-Host)
        IF ($Response -eq "Y") {

            Set-AudioDevice -PlayBackVolume $AudioOutputVolume
            $PlaybackVolume = Get-AudioDevice -PlaybackVolume
            Write-Host "PlaybackVolume ="$PlaybackVolume -ForeGroundColor Green

        }

    }

}




Function AudioInputSettings {

    ### Check Audio Input Settings
    Write-Host "`nChecking Audio Input Settings..."
    # Check if the desired device is set as the default recording device
    $RecordingDevice = Get-AudioDevice -list | Where-Object {$_.Name -eq $Microphone} | Select-Object Index, Name, Default, DefaultCommunication
    $RecordingDevice | Select-Object Name, Default, DefaultCommunication | Format-List
    IF ($RecordingDevice.Default -match "False") {

        $Response = $(Write-Host $Microphone "is not the default recording device. Set this as default? Y/N: " -ForegroundColor Yellow -NoNewLine; Read-Host)
        IF ($Response -eq "Y") {

            Set-AudioDevice $RecordingDevice.Index -DefaultOnly | Out-Null
            Get-AudioDevice -index $RecordingDevice.Index | Select-Object Name, Default, DefaultCommunication | Format-List

        }

    } 
    # Check if the desired device is set as the default recording communication device
    IF ($RecordingDevice.DefaultCommunication -match "False") {

        $Response = $(Write-Host $Microphone "is not the default recording communication device. Set this as default? Y/N: " -ForegroundColor Yellow -NoNewLine; Read-Host)
        IF ($Response -eq "Y") {

            Set-AudioDevice $RecordingDevice.Index -CommunicationOnly | Out-Null
            Get-AudioDevice -index $RecordingDevice.Index | Select-Object Name, Default, DefaultCommunication | Format-List

        } 

    }
    # Check microphone volume
    $RecordingVolume = Get-AudioDevice -RecordingVolume
    IF ($RecordingVolume.Trim("%") -eq $MicrophoneVolume) {

        Write-Host "RecordingVolume ="$RecordingVolume -ForeGroundColor Green

    } ELSE {

        $Response = $(Write-Host "MicrophoneVolume ="$RecordingVolume"`nSet MicrophoneVolume to" $MicrophoneVolume"%? Y/N: " -ForegroundColor Yellow -NoNewLine; Read-Host)
        IF ($Response -eq "Y") {

            Set-AudioDevice -RecordingVolume $MicrophoneVolume
            $RecordingVolume = Get-AudioDevice -RecordingVolume
            Write-Host "RecordingVolume ="$RecordingVolume -ForeGroundColor Green

        }

    }

}


Function MouseAcceleration {
    
    ### Check Mouse Acceleration Settings
    Write-Host "`nChecking Mouse Acceleration Setting..."
    $RegKey = "HKCU:\Control Panel\Mouse"
    $MouseSpeed = (Get-ItemProperty -Path $RegKey -Name MouseSpeed).MouseSpeed # MouseSpeed = MouseAcceleration
    IF ($MouseSpeed -eq $MouseAcceleration) {
        
        IF ($MouseSpeed -eq "1") {

            Write-Host "Mouse Acceleration = ON" -ForeGroundColor Green

        } ELSEIF ($MouseSpeed = "0") {

            Write-Host "Mouse Acceleration = OFF" -ForeGroundColor Green
            
        }

    } ELSEIF ($MouseSpeed -eq "0") {

        $Response = $(Write-Host "Mouse Acceleration = OFF.`nSet to ON? Y/N: " -ForeGroundColor Yellow -NoNewLine; Read-Host)
        IF ($Response -eq "Y") {

            Set-ItemProperty -Path $RegKey -Name MouseSpeed -Value 1
            Set-ItemProperty -Path $RegKey -Name MouseThreshold1 -Value 6
            Set-ItemProperty -Path $RegKey -Name MouseThreshold2 -Value 10
            Write-Host "Mouse Acceleration switched ON. Change will take effect on next restart." -ForeGroundColor Green

        }

    } ELSEIF ($MouseSpeed -eq "1") {

        $Response = $(Write-Host "Mouse Acceleration = ON.`nSet to OFF? Y/N: " -ForeGroundColor Yellow -NoNewLine; Read-Host)
        IF ($Response -eq "Y") {

            Set-ItemProperty -Path $RegKey -Name MouseSpeed -Value 0
            Set-ItemProperty -Path $RegKey -Name MouseThreshold1 -Value 0
            Set-ItemProperty -Path $RegKey -Name MouseThreshold2 -Value 0
            Write-Host "Mouse Acceleration switched OFF. Change will take effect on next restart." -ForeGroundColor Green

        }

    }

}




Function MousePointerSpeed {

    ### Check Mouse Pointer Speed Settings
    Write-Host "`nChecking Mouse Pointer Speed..."
    $RegKey = "HKCU:\Control Panel\Mouse"
    $MouseSensitivity = (Get-ItemProperty -Path $RegKey -Name MouseSensitivity).MouseSensitivity # MouseSensitivity = Windows Mouse Pointer Speed

    IF ($MouseSensitivity -eq $MousePointerSpeed) {
        
        Write-Host "Mouse Pointer Speed = $MouseSensitivity" -ForeGroundColor Green

    } ELSE {

        $Response = $(Write-Host "Mouse Pointer Speed =" $MouseSensitivity".`nSet to "$MousePointerSpeed"? Y/N: " -ForeGroundColor Yellow -NoNewLine; Read-Host)
        IF ($Response -eq "Y") {

            Set-ItemProperty -Path $RegKey -Name MouseSensitivity -Value $MousePointerSpeed
            Write-Host "Mouse Pointer Speed set to $MousePointerSpeed. Change will take effect on next restart." -ForeGroundColor Green

        }

    }

}




#Call the functions:
AudioDeviceCmdlets
AudioInputSettings
AudioOutputSettings
MouseAcceleration
MousePointerSpeed

Write-Output "`nChecks complete. Press any key to close..."
$Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown") | Out-Null
