# Path to your audio file
$audioFile = "C:\Windows\Media\Ring01.wav"

# Check if the audio file exists
if (!(Test-Path -Path $audioFile)) {
    Write-Host "Audio file not found. Please check the path."
    exit
}

# Path to Windows Media Player
$wmplayerPath = "C:\Program Files (x86)\Windows Media Player\wmplayer.exe"

# Function to play the audio file using Windows Media Player minimized
function Play-Audio {
    $player = Start-Process -FilePath $wmplayerPath -ArgumentList "`"$audioFile`"" -PassThru

    # Looping mechanism with wait time
    while ($true) {
        Start-Sleep -Seconds 6
        if (-not (Get-Process -Name wmplayer -ErrorAction SilentlyContinue)) {
            taskkill /f /IM NoiseMaker.exe
            exit
        }

        Stop-Process -Name wmplayer -Force
        Start-Sleep -Seconds 1
        $player = Start-Process -FilePath $wmplayerPath -ArgumentList "`"$audioFile`"" -PassThru
    }
}

Play-Audio
