#Stage 1: Get the Administrator Role.

if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }


#Stage 2: Create a Temp folder, download the required files and put them in the Temp Folder

    #Create the Temp Folder
        New-Item -ItemType Directory -Path c:\office
    
    # Download Office Zip, which contains:
    # - The Official Office Setup
    # - The Configuration file, with the basic Office Apps (Word, Access, Visio, Excel, PowerPoint)
    # - The Activation Script (No virus, No Tricks) 
        # !!! The script is just an automated version of the official MAS Activation Repo.
        
        Invoke-WebRequest "https://raw.githubusercontent.com/Eklip5e/Office-Suite/main/office.zip" -OutFile c:\office.zip
    
        Expand-Archive C:\office.zip -DestinationPath C:\office

Start-Sleep -Seconds 2
    
    #Start the Installation
        & "c:\office\setup.exe" /configure "c:\office\conf.xml"

#Stage 3: Check if the Installation has finished, if so, start the activation script.
    Start-Sleep -Seconds 5
    $a = 1
    DO
    {
        
        if((get-process "OfficeC2RClient" -ea SilentlyContinue) -eq $Null)
        { 
            Stop-Process -force -processname "OfficeClickToRun"
            if((get-process "OfficeClickToRun" -ea SilentlyContinue) -eq $Null)
            { 
            $a = 2
            "Office Has Finished"
            & "c:\office\Activate.cmd"
            break
            }
        }
    }
    WHILE ($a = 1)

#Stage 4: Everything has finished and Office is installed and activated, delete all the temp files.

"Cleaning up some mess..."

    #Check if the folder is present, if so, delete it and make sure it's gone.
if (Test-Path "C:\office" -IsValid) 
{
    Remove-Item –path c:\office –recurse
    Remove-Item –path c:\office.zip –recurse
    "Done"
     Break
}
else {
    "Done."
     Break
     }




