This is a command line tool for Windows, for setting default audio devices (input/output, console/communication) by finding a supplied device name.

Command line help:

Set the default audio device.

        Set_audio_device.exe [-in/-out] (-cons/-comm) ["Device Name"]

Must specify audio direction:
        input(-in) or output(-out)
Optionally specify the audio role:
        console/general purpose(-cons) or communication(-comm)
        If ommitted, both are set
Lastly specify the device name:
        Case sensitive.
        Use quotes if spaces are necessary.
        Can be a substring of the device name.
        If it matches multiple devices, last matching device will be set.

Example usage:

        Set_audio_device.exe -in "Microphone" -out -comm "Headset" -out -cons "Speakers"


List available device names:

        Set_audio_device.exe (-in/-out) -list


The solution uses Visual Studio 2019 and VC++17