#Requires -Version 5
# Notes:
# - Using -Verbose or -Debug when calling the script will not kill the PowerShell process after the MainWindow closes

[CmdletBinding()]
Param()

# Import the MaterialDesign module. Exit if it fails to load
TRY {
    Import-Module "$PSScriptRoot\MaterialDesign\MaterialDesignColors.dll" -ErrorAction Stop
    Import-Module "$PSScriptRoot\MaterialDesign\MaterialDesignThemes.Wpf.dll" -ErrorAction Stop
}
CATCH {
    Write-Error $_.Exception | Format-List -Force
    Write-Host 'Press any key to continue ...'
    [void]$host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Add the WPF library
Add-Type -AssemblyName PresentationFramework

#############################################################################################
#
# SHARED VARIABLES AND FUNCTIONS
#
#############################################################################################

# Set up synchronized hashtables to:
# - Run the UI
# - Share data across runspaces
$UI = [hashtable]::Synchronized(@{})
$SharedData = [hashtable]::Synchronized(@{})

# The path to the script, to pass into the UI runspace, so it can find MainWindow.xml
$Script_Path = $PSScriptRoot

$SharedData.CloseWindow = $false

$Shared_Variables = @(
    # The synchronized hashtables
    'SharedData'
    'UI'

    # The runspaces
    #'UI_Runspace'
    #'Worker_Runspace'

    # Shared functions
    'Shared_Functions'

    # Display preferences
    'DebugPreference'
    'ErrorActionPreference'
    'InformationPreference'
    'ProgressPreference'
    'VerbosePreference'
    'WarningPreference'

    # Misc variables
    'Script_Path'
)

# All shared functions are declared here, so they can be imported into all the runspaces
$Shared_Functions = [scriptblock]::Create({


})

#############################################################################################
#
# RUNSPACE SETUP
#
#############################################################################################

# Create a runspace (separate PowerShell thread) to perform the work from the application
# This ensures modules are only loaded once and keeps PSSessions that have been established (eg. to Office365)
$Worker_Runspace = [runspacefactory]::CreateRunspace()
$Worker_Runspace.ApartmentState = 'STA'
$Worker_Runspace.ThreadOptions = 'ReuseThread'
$Worker_Runspace.Name = 'Worker_Runspace'
$Worker_Runspace.Open()

$Worker_Instance = [powershell]::Create()
$Worker_Instance.Runspace = $Worker_Runspace

# Create a runspace to run the User Interface
# This will keep the UI responsive while the Worker runspace is executing
$UI_Runspace = [runspacefactory]::CreateRunspace()
$UI_Runspace.ApartmentState = 'STA'
$UI_Runspace.ThreadOptions = 'ReuseThread'
$UI_Runspace.Name = 'UI_Runspace'
$UI_Runspace.Open()

$UI_Instance = [powershell]::Create()
$UI_Instance.Runspace = $UI_Runspace

# Add all the shared variables into the runspaces
@('Worker_Runspace','UI_Runspace') | ForEach-Object {
    foreach ($Shared_Variable in $Shared_Variables) {
        Write-Verbose ('Adding variable "{0}" to runspace: {1}' -f $Shared_Variable,$_)
        (Get-Variable -Name $_ -ValueOnly).SessionStateProxy.SetVariable($Shared_Variable,(Get-Variable -Name $Shared_Variable -ValueOnly))
    }
}


#############################################################################################
#
# WORKER RUNSPACE CODE
#
#############################################################################################

# Add all the shared variables into the runspace
foreach ($Shared_Variable in $Shared_Variables) {
    $UI_Runspace.SessionStateProxy.SetVariable($Shared_Variable,(Get-Variable -Name $Shared_Variable -ValueOnly))
}

# All the code that will be run only in the worker runspace
$Worker_Code = {
    # Source all the shared functions
    . $Shared_Functions
}

# Launch the Worker runspace
[void]$Worker_Instance.AddScript($Worker_Code)
$Worker_Job = $Worker_Instance.BeginInvoke()


#############################################################################################
#
# UI RUNSPACE CODE
#
#############################################################################################

# All the code that will be run in the UI runspace
$UI_Code = {
    # Source all the shared functions
    . $Shared_Functions

    #########################################################################################
    # Load the XAML and set up the MainWindow
    #########################################################################################
    # Parts taken from:
    #    https://gallery.technet.microsoft.com/Exchange-Log-Level-GUI-f9e8cb21
    #    http://blogs.technet.com/b/heyscriptingguy/archive/2014/08/01/i-39-ve-got-a-powershell-secret-adding-a-gui-to-scripts.aspx
    #    http://foxdeploy.com/2015/04/16/part-ii-deploying-powershell-guis-in-minutes-using-visual-studio/

    $WindowDefinition = Get-Content "$Script_Path\MainWindow.xaml"
    # Filter out some elements that are not needed
    [xml]$WindowXAML = $WindowDefinition -replace '^<Win.*','<Window' -replace 'mc:Ignorable="d"',''
    $XMLReader = (New-Object System.Xml.XmlNodeReader $WindowXAML)
    $UI.Window = [Windows.Markup.XamlReader]::Load($XMLReader)
    # Generate variables from all form objects
    $WindowXAML.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | ForEach-Object {
        $UI.$($_.Name) = $UI.Window.FindName($_.Name)
    }

    #########################################################################################
    # UI Events
    #########################################################################################
    $UI.Window.Add_Closed({
        $SharedData.CloseWindow = $true
    })

    #########################################################################################
    # UI Actions
    #########################################################################################


    #----------------------------------------------------------------------------------------
    # LAUNCH THE WINDOW - must be last
    #----------------------------------------------------------------------------------------
    [void]$UI.Window.ShowDialog()
}

# Launch the User Interface inside the runspace
[void]$UI_Instance.AddScript($UI_Code)
$UI_Job = $UI_Instance.BeginInvoke()

#############################################################################################
#
# Debugging and Cleanup
#
#############################################################################################

while (-not $SharedData.CloseWindow) {
    # Debugging
    if ($UI_Job -or $Worker_Job) {
        # Loop through each of the instances, and through each stream from those instances, and return the values from the stream before clearing
        foreach ($Instance_Name in ('UI_Instance','Worker_Instance')) {
            $Instance_Streams = (Get-Variable -Name $Instance_Name -ValueOnly).Streams
            foreach ($Instance_Stream_Name in ($Instance_Streams | Get-Member -MemberType Property).Name) {
                if ($Instance_Streams.$Instance_Stream_Name) {
                    Write-Verbose ('Start of: {0} stream from {1}' -f $Instance_Stream_Name,$Instance_Name)
                    $Instance_Streams.$Instance_Stream_Name
                    $Instance_Streams.$Instance_Stream_Name.Clear()
                    Write-Verbose ('End of: {0} stream from {1}' -f $Instance_Stream_Name,$Instance_Name)
                }
            }
        }
    }
}

if ($SharedData.CloseWindow) {
    Write-Verbose "Killing worker runspace"
    $Worker_Instance.EndInvoke($Worker_Job)
    $Worker_Runspace.Close()
    $Worker_Instance.Dispose()

    Write-Verbose "Killing UI runspace"
    $UI_Instance.EndInvoke($UI_Job)
    $UI_Runspace.Close()
    $UI_Instance.Dispose()

    Write-Verbose "Garbage collection"
    [system.gc]::Collect()
    [system.gc]::WaitForPendingFinalizers()
    if ((-not $PSBoundParameters['Debug']) -and (-not $PSBoundParameters['Verbose'])) { Stop-Process $PID }
}
