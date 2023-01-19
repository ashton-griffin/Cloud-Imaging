
#exported functions#
function Invoke-NetworkTestingClient {
    [CmdletBinding()]
    param(
        [switch]$DebugWithGlobalVariables
    )
    try {
        (Get-Runspace -Name MainWindow).Dispose()
        (Get-Runspace -Name Splash).Dispose()
    } catch {
        # No previous runspace to remove
    }

    if ($DebugWithGlobalVariables) {
        $Global:uiHash = [hashtable]::Synchronized(@{})
        $Global:variableHash = [hashtable]::Synchronized(@{})
    } else {
        $uiHash = [hashtable]::Synchronized(@{})
        $variableHash = [hashtable]::Synchronized(@{})
    }
    
    #store runspaces in jobs array so we can dispose of them when we're done
    $Jobs = @{}
    $uiHash.Jobs = $Jobs
    $Hosts = @{}
    $uiHash.Hosts = $Hosts
    $Handles = @{}
    $uiHash.Handles = $Handles
    $RunspaceOutput = @{}
    $uiHash.RunspaceOutput = $RunspaceOutput    
    
    $uiHash.resultsHash = $null
    $uiHash.AdminMode = $false
    
    #set static paths for info links
    $variableHash.navInfoOs = "https://docs.microsoft.com/en-us/microsoftteams/hardware-requirements-for-the-teams-app"
    $variableHash.navInfoInternet = "https://docs.microsoft.com/en-us/microsoftteams/prepare-network"
    $variableHash.navInfoHeadset = "https://docs.microsoft.com/en-us/microsoftteams/devices/usb-devices"
    $variableHash.navInfoTool = "https://docs.microsoft.com/en-us/microsoftteams/3-envision-evaluate-my-environment#network-remediation"
    $variableHash.navCallQuality = "https://docs.microsoft.com/en-ca/SkypeForBusiness/optimizing-your-network/media-quality-and-network-connectivity-performance"
    $variableHash.InternetTestIp = "13.107.64.2" #need this as we may not have the tool installed yet to pull the .config file and read the setting
    $variableHash.ReadyToStartTests = $false

    $variableHash.ThisModule = Get-Module NetworkTestingCompanion
    [array]$variableHash.allModules = Get-Module NetworkTestingCompanion -ListAvailable
    $variableHash.RootPath = Split-Path -Path $variableHash.ThisModule.Path
    $variableHash.osMinVer = "6.1.7601"
    
    $variableHash.resultsAnalyzerTempFile = $env:TEMP + "\results_temp.csv"
    $variableHash.resultsAnalyzerTextFile = $env:TEMP + "\results_analyzer.txt"
    $variableHash.connectivityCheckResults = $env:TEMP + "\connectivity_results.txt"

    [System.Version]$variableHash.ToolVersion = "2018.22.1.9"
    $variableHash.toolName = "Microsoft Skype for Business Network Assessment Tool"
    $variableHash.networkAssessmentToolPath = "${env:ProgramFiles(x86)}\Microsoft Skype for Business Network Assessment Tool\NetworkAssessmentTool.exe"
    $variableHash.networkAssessmentToolConfig = "${env:ProgramFiles(x86)}\Microsoft Skype for Business Network Assessment Tool\NetworkAssessmentTool.exe.config"
    $variableHash.networkAssessmentToolWorkingDirectory = "${env:ProgramFiles(x86)}\Microsoft Skype for Business Network Assessment Tool\"

    $uiHash.checkMark = $variableHash.RootPath + "\assets\check.png"
    $uiHash.errorMark = $variableHash.RootPath + "\assets\error.png"
    $uiHash.warningMark = $variableHash.RootPath + "\assets\warning.png"
    $uiHash.questionMark = $variableHash.RootPath + "\assets\question.png"

    $uiHash.AutoUpdate = (Get-ToolRegistryEntries).AutoUpdate
    $uiHash.CheckUpdateContent = "Check for Updates"

    $splashMain = {
        $uiHash.Hosts.("RspSplash") = $Host
        Add-Type -AssemblyName PresentationFramework
        
        $uiContent = Get-Content -Path ($variableHash.rootPath + "\Splash.xaml")
        
        [xml]$xAML = $uiContent -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
        $xmlReader = (New-Object System.Xml.XmlNodeReader $xAML)
        $uiHash.Splash = [Windows.Markup.XamlReader]::Load($xmlReader)

        $xAML.SelectNodes("//*[@Name]") | ForEach-Object {

            $uiHash.Add($_.Name, $uiHash.Splash.FindName($_.Name))

        }

        $uiHash.Splash.Add_SourceInitialized(
            {

                $assets = $variableHash.RootPath + "\assets\"
                $uiHash.imgSplashLogo.Source = $variableHash.RootPath + "\assets\corplogo_small.png"
                $uiHash.imgTeamsLogo.Source = $variableHash.RootPath + "\assets\msft_teams_logo_small.png"
                $uiHash.imgSFBLogo.Source = $variableHash.RootPath + "\assets\sfb_logo_small.png"
                $uiHash.Splash.Icon = $assets + "msft_logo.png"

                $uiUpdateSplash = {

                    $uiHash.txtSplashStatus.Text = $uiHash.SplashStatusText

                }

                #Create timer to handle updating the UI
                $timer = new-object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]"0:0:0:0.100"
                $timer.Add_Tick($uiUpdateSplash)
                $timer.Start()
            }
        )

        $uiHash.Splash.Add_Loaded(
            {

                $wid = [System.Security.Principal.WindowsIdentity]::GetCurrent()
                $prp = New-Object System.Security.Principal.WindowsPrincipal($wid)
                $adm = [System.Security.Principal.WindowsBuiltInRole]::Administrator
                
                if ($prp.IsInRole($adm)) {
                    
                    $uiHash.AdminMode = $true

                }

                $codeBootstrap = {
                    
                    Invoke-NetworkTestingCompanionAutoUpdate -CheckForUpdates
                
                }

                Invoke-NewRunspace -codeBlock $codeBootstrap -RunspaceHandleName Bootstrap

            }
        )

        $uiHash.Splash.ShowDialog() | Out-Null
        $uiHash.Splash.Error = $Error

    }

    Write-Output "Attempting to start the UI..."
    Invoke-NewRunspace -codeBlock $splashMain -RunspaceHandleName Splash

    do {

        Start-Sleep -Milliseconds 500

    } until ($uiHash.Handles.Bootstrap.IsCompleted)

    #close out possible runspaces

    if ($variableHash.ExitApplication) {

        Write-Output "Exiting due to new update..."
        exit

    }
    
    $mainWindow = {

        $uiHash.Hosts.("RspMain") = $Host
        Add-Type -AssemblyName PresentationFramework
        
        $uiContent = Get-Content -Path ($variableHash.RootPath + "\MainWindow.xaml")
        
        [xml]$xAML = $uiContent -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
        $xmlReader = (New-Object System.Xml.XmlNodeReader $xAML)
        $uiHash.Window = [Windows.Markup.XamlReader]::Load($xmlReader)

        $xAML.SelectNodes("//*[@Name]") | ForEach-Object {

            $uiHash.Add($_.Name, $uiHash.Window.FindName($_.Name))

        }

        #region EVENTS #
            $uiHash.Window.Add_SourceInitialized(
                {

                    $uiHash.Window.Resources.ConnectivityTimeout = @(90,120,150,190)
                    $uiHash.Window.Resources.NumAudioTests = @(1..50)
                    $uiHash.Window.Resources.AudioTestDelay = @(1,5,10,20,30,60,90,120)

                    [int]$uiHash.NumAudioTests = (Get-ToolRegistryEntries).NumberOfAudioTestIterations
                    [int]$uiHash.AudioTestDelay = (Get-ToolRegistryEntries).TestIntervalInSeconds
                    [int]$uiHash.ConnectivityTimeout = (Get-ToolRegistryEntries).ConnectivityTimeoutInSeconds

                    if ($uiHash.NumAudioTests) {

                        $uiHash.cmbNumAudioTests.SelectedItem = $uiHash.NumAudioTests

                    } else {

                        $uiHash.NumAudioTests = $uiHash.cmbNumAudioTests.SelectedItem = $uiHash.Window.Resources.NumAudioTests[0]

                    }

                    if ($uiHash.AudioTestDelay) {

                        $uiHash.cmbAudioTestDelay.SelectedItem = $uiHash.AudioTestDelay

                    } else {

                        $uiHash.AudioTestDelay = $uiHash.cmbAudioTestDelay.SelectedItem = $uiHash.Window.Resources.AudioTestDelay[0]

                    }

                    if ($uiHash.ConnectivityTimeout) {
                        
                        $uiHash.cmbConnectivityTimeout.SelectedItem = $uiHash.ConnectivityTimeout

                    } else {

                        $uiHash.ConnectivityTimeout = $uiHash.cmbConnectivityTimeout.SelectedItem = $uiHash.Window.Resources.ConnectivityTimeout[0]

                    }

                    if ($uiHash.AdminMode){

                        $uiHash.Window.Title = "Administrator: " + $uiHash.Window.Title

                    }
                   
                    #set image locations
                    $assets = $variableHash.RootPath + "\assets\"

                    $uiHash.imgMSFTLogo.Source = $assets + "msft_logo.png"
                    $uiHash.imgInternetLogo.Source = $assets + "network.png"
                    $uiHash.imgHeadsetLogo.Source = $assets + "headset.png"
                    $uiHash.imgToolLogo.Source = $assets + "tool.png"
                    $uiHash.Window.Icon = $assets + "msft_logo.png"

                    $uiHash.TestQualitySource = $uiHash.questionMark
                    $uiHash.TestQualityDetailSource = $uiHash.questionMark
                    $uiHash.TestConnectivitySource = $uiHash.questionMark
                    $uiHash.TestConnectivityDetailSource = $uiHash.questionMark
                    
                    $uiHash.PacketLossRateSource = $uiHash.questionMark
                    $uiHash.RoundTripTimeSource = $uiHash.questionMark
                    $uiHash.JitterSource = $uiHash.questionMark
                    $uiHash.PacketReorderRatioSource = $uiHash.questionMark

                    $uiHash.imgInfoOs.Source = $assets + "info.png"
                    $uiHash.imgInfoInternet.Source = $assets + "info.png"
                    $uiHash.imgInfoHeadset.Source = $assets + "info.png"
                    $uiHash.imgInfoTool.Source = $assets + "info.png"
                    $uiHash.imgEmailResults.Source = $assets + "email.png"
                    $uiHash.reportQuality.Source = $assets + "doc.png"
                    $uiHash.reportConnectivity.Source = $assets + "doc.png"

                    $uiHash.ActionButtonText = "Checking"
                    $uiHash.ActionButtonFill = "#FFFBBC0B" #yellow
                    $uiHash.CheckUpdateEnabled = $true

                    $uiUpdateBlock = {

                        $uiHash.barTest.Value = $uiHash.ProgressValue

                        $uiHash.txtStatus.Text = $uiHash.StatusText
                        $uiHash.txtActionButton.Text = $uiHash.ActionButtonText
                        $uiHash.elipActionButton.Fill = $uiHash.ActionButtonFill
                        
                        $uiHash.txtConnectivityDetail.Text = $uiHash.ConnectivityDetailText
                        $uiHash.txtPacketLossRate.Text = $uiHash.PacketLossRateText
                        $uiHash.txtRoundTripTime.Text = $uiHash.RoundTripTimeText
                        $uiHash.txtJitter.Text = $uiHash.JitterText
                        $uiHash.txtPacketReorderRatio.Text = $uiHash.PacketReorderRatioText

                        #this block needs to be set as we look for the update from a previous runspace
                        $uiHash.txtCurrentVersion.Text = $uiHash.CurrentVersionText
                        $uiHash.txtPSGVersion.Text = $uiHash.PSGVersionText

                        #update check/x-marks
                        $uiHash.imgTestQuality.Source = $uiHash.TestQualitySource
                        $uiHash.imgTestConnectivity.Source = $uiHash.TestConnectivitySource
                        $uiHash.imgQualityDetail.Source = $uiHash.TestQualityDetailSource
                        $uiHash.imgConnectivityDetail.Source = $uiHash.TestConnectivityDetailSource
                        $uiHash.imgPacketLossRate.Source = $uiHash.PacketLossRateSource
                        $uiHash.imgRoundTripTime.Source = $uiHash.RoundTripTimeSource
                        $uiHash.imgJitter.Source = $uiHash.JitterSource
                        $uiHash.imgPacketReorderRatio.Source = $uiHash.PacketReorderRatioSource

                        #validate pre-requisites
                        if ($variableHash.OSReady -and $variableHash.InternetReady -and $variableHash.ToolReady -and $variableHash.TestingIdle) {
                            $uiHash.ActionButtonFill = "#FF80CC28"
                            $uiHash.ActionButtonText = "Start"
                        }
                    }

                    #Create timer to handle updating the UI
                    $timer = new-object System.Windows.Threading.DispatcherTimer
                    $timer.Interval = [TimeSpan]"0:0:0:0.30"
                    $timer.Add_Tick($uiUpdateBlock)
                    $timer.Start()
                }
            )

            $uiHash.Window.Add_Loaded(
                {
                    #close splash screen
                    $uiHash.Splash.Dispatcher.Invoke([action]{$uiHash.Splash.Close()})
                    (Get-Runspace -Name Bootstrap).Dispose()
                    (Get-Runspace -Name Splash).Dispose()

                    function CheckOs {
                    
                        $osVersion = [System.Environment]::OSVersion

                        if ($osVersion.Version -ge $variableHash.osMinVer) {

                            if ([System.Environment]::Is64BitOperatingSystem) {

                                $uiHash.txtOs.Text = "Windows version $($osVersion.Version.ToString()) has been detected and is a 64-bit operating system."
                                $uiHash.imgOs.Source = $uiHash.checkMark
                                $variableHash.OSReady = $true

                            } else {

                                #not 64-bit
                                $uiHash.txtOs.Text = "Windows version $($osVersion.Version.ToString()) has been detected and is not a 64-bit operating system."
                                $uiHash.imgOs.Source = $uiHash.errorMark

                            }
                        } else {

                            #not Win7 or higher
                            $uiHash.txtOs.Text = "Windows version $($osVersion.Version.ToString()) has been detected and does not match the minimum required version. Click the info button to learn more."
                            $uiHash.imgOs.Source = $uiHash.errorMark

                        }

                    }

                    CheckOs

                    $preReqInternetCheck = {

                        Invoke-InternetConnectionTest

                    }                    

                    $preReqEndpointCheck = {

                        Invoke-CheckCertifiedEndpoint

                    }
                        
                    $preReqToolCheck = {

                        Invoke-CheckToolInstall

                    }
                        
                    #these two are used elsewhere in the module and they take longer so we
                    #put them in a runspace to surface the UI faster and keep them in root functions
                    #so other functions can call them

                    Invoke-NewRunspace -codeBlock $preReqInternetCheck -RunspaceHandleName PreReqInternetCheck
                    Invoke-NewRunspace -codeBlock $preReqEndpointCheck -RunspaceHandleName PreReqEndpointCheck
                    Invoke-NewRunspace -codeBlock $preReqToolCheck -RunspaceHandleName PreReqToolCheck
                    
                    Invoke-CalculateEstimatedTime
                    
                    $uiHash.StatusText = "Click the start button to begin testing"
                    $variableHash.TestingIdle = $true

                }
            )

            $uiHash.Window.Add_Closing(
                {
                    #this is where we do our cleanup to prevent memory leaks.
                    #we don't want these runspaces to consume memory if they're
                    #not in use even after the window has been closed.

                    Remove-Item $variableHash.resultsAnalyzerTextFile -Force -ErrorAction SilentlyContinue
                    Remove-Item $variableHash.connectivityCheckResults -Force -ErrorAction SilentlyContinue
                    Remove-Item $variableHash.resultsAnalyzerTempFile -Force -ErrorAction SilentlyContinue
                    
                    Set-ToolRegistryEntries

                }
            )

            $uiHash.Window.Add_Closed(
                {
                    $uiHash.Jobs.GetEnumerator() | ForEach-Object {

                        if ($_.Name -eq "MainWindow") {

                            #leave MainWindow as this causes issues. We clean it up on launch. Maybe a better way?

                        } else {

                            (Get-Runspace -Name $_.Name).Dispose()

                        }
                    }
                }
            )

            $uiHash.cmbConnectivityTimeout.Add_DropDownClosed(
                {

                    $uiHash.ConnectivityTimeout = $uiHash.cmbConnectivityTimeout.SelectedItem

                }
            )

            $uiHash.cmbNumAudioTests.Add_DropDownClosed(
                {

                    [int]$uiHash.NumAudioTests = $uiHash.cmbNumAudioTests.SelectedItem
                    Invoke-CalculateEstimatedTime
                }
            )

            $uiHash.cmbAudioTestDelay.Add_DropDownClosed(
                {

                    [int]$uiHash.AudioTestDelay = $uiHash.cmbAudioTestDelay.SelectedItem
                    Invoke-CalculateEstimatedTime
                }
            )

            $uiHash.barTest.Add_ValueChanged(
                {

                    $uiHash.Window.Resources.ProgressValue = $uiHash.barTest.Value / 100

                }
            )

            $uiHash.elipActionButton.Add_MouseUp(
                {
                    Invoke-ActionButton
                }
            )
            
            $uiHash.txtActionButton.Add_MouseUp(
                {
                    Invoke-ActionButton
                }
            )

            $uiHash.navMailTo.Add_Click(
                {
                    Start-Process $uiHash.navMailTo.NavigateUri
                }
            )

            $uiHash.navMyAdvisor.Add_Click(
                {
                    Start-Process $uiHash.navMyAdvisor.NavigateUri
                }
            )

            $uiHash.imgInfoOs.Add_MouseUp(
                {
                    Start-Process $variableHash.navInfoOs
                }
            )

            $uiHash.imgInfoInternet.Add_MouseUp(
                {
                    Start-Process $variableHash.navInfoInternet
                }
            )

            $uiHash.imgInfoHeadset.Add_MouseUp(
                {
                    Start-Process $variableHash.navInfoHeadset
                }
            )

            $uiHash.imgInfoTool.Add_MouseUp(
                {
                    Start-Process $variableHash.navInfoTool
                }
            )

            $uiHash.reportQuality.Add_MouseUp(
                {
                    Start-Process $variableHash.resultsAnalyzerTextFile
                }
            )

            $uiHash.reportConnectivity.Add_MouseUp(
                {
                    Start-Process $variableHash.connectivityCheckResults
                }
            )

            $uiHash.imgTestQuality.Add_MouseUp(
                {
                    Start-Process $variableHash.navCallQuality
                }
            )

            $uiHash.imgPacketLossRate.Add_MouseUp(
                {
                    Start-Process $variableHash.navCallQuality
                }
            )
            
            $uiHash.imgRoundTripTime.Add_MouseUp(
                {
                    Start-Process $variableHash.navCallQuality
                }
            )
            
            $uiHash.imgJitter.Add_MouseUp(
                {
                    Start-Process $variableHash.navCallQuality
                }
            )

            $uiHash.imgPacketReorderRatio.Add_MouseUp(
                {
                    Start-Process $variableHash.navCallQuality
                }
            )

            $uiHash.imgQualityDetail.Add_MouseUp(
                {
                    Start-Process $variableHash.navCallQuality
                }
            )

        #end region

        [void]$uiHash.Window.Dispatcher.InvokeAsync{$uiHash.Window.ShowDialog()}.Wait()
        $uiHash.Error = $Error
    }

    Invoke-NewRunspace -codeBlock $mainWindow -RunspaceHandleName MainWindow

    if (!$DebugWithGlobalVariables) {

        while (!$uiHash.Handles.MainWindow.IsCompleted) {

            Start-Sleep -Seconds 1

        }
    
        Write-Output "Closing the application..."

    }

}

function Invoke-NewRunspace {
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)][scriptblock]$codeBlock,
        [parameter(Mandatory=$true)][string]$RunspaceHandleName
    )

    $testingRunspace = [runspacefactory]::CreateRunspace()
    $testingRunspace.ApartmentState = "STA"
    $testingRunspace.ThreadOptions = "ReuseThread"
    $testingRunspace.Open()
    $testingRunspace.SessionStateProxy.SetVariable("uiHash",$uiHash)
    $testingRunspace.SessionStateProxy.SetVariable("variableHash",$variableHash)
    $testingRunspace.Name = $RunspaceHandleName

    $uiHash.RunspaceOutput.($RunspaceHandleName) = New-Object System.Management.Automation.PSDataCollection[psobject]

    $testingCmd = [PowerShell]::Create().AddScript($codeBlock)
    
    $testingCmd.Runspace = $testingRunspace
    
    $testingHandle = $testingCmd.BeginInvoke($uiHash.RunspaceOutput.($RunspaceHandleName),$uiHash.RunspaceOutput.($RunspaceHandleName))

    #store the handle in the global sync'd hashtable; arraylist we'll use the window.closed() event to clean up
    $uiHash.Jobs.($RunspaceHandleName) = $testingCmd
    $uiHash.Handles.($RunspaceHandleName) = $testingHandle
}

function Invoke-ActionButton {

    $actionBlock = {

        if ($uiHash.ActionButtonText -eq "Start") {

            Invoke-InternetConnectionTest
            if (!$variableHash.InternetReady){
                throw
            }

            Remove-Item $variableHash.resultsAnalyzerTextFile -Force -ErrorAction SilentlyContinue
            Remove-Item $variableHash.connectivityCheckResults -Force -ErrorAction SilentlyContinue
            Remove-Item $variableHash.resultsAnalyzerTempFile -Force -ErrorAction SilentlyContinue
            Remove-Item $variableHash.AudioTestSavePath -Force -ErrorAction SilentlyContinue
            Remove-Item $variableHash.ConnectionTestSavePath -Force -ErrorAction SilentlyContinue

            $variableHash.TestingIdle = $false
            $uiHash.TestQualitySource = $uiHash.questionMark
            $uiHash.TestQualityDetailSource = $uiHash.questionMark
            $uiHash.TestConnectivitySource = $uiHash.questionMark
            $uiHash.TestConnectivityDetailSource = $uiHash.questionMark
            $uiHash.PacketLossRateSource = $uiHash.questionMark
            $uiHash.RoundTripTimeSource = $uiHash.questionMark
            $uiHash.JitterSource = $uiHash.questionMark
            $uiHash.PacketReorderRatioSource = $uiHash.questionMark

            $uiHash.JitterText = ""
            $uiHash.PacketLossRateText = ""
            $uiHash.PacketReorderRatioText = ""
            $uiHash.RoundTripTimeText = ""

            $totalIterations = $uiHash.NumAudioTests

            $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.barTest.Visibility = "Visible"})

            Invoke-ConnectivityTest

            if ($variableHash.NumTestsFromXml -ne "1") {

                $uiHash.StatusText = "Config file NumIterations is {0}. Please change it back to 1." -f $variableHash.NumTestsFromXml
                throw

            }

            for ($testIteration = 1; $testIteration -le $totalIterations; $testIteration++) {

                if ($uiHash.StopTesting) {

                    break;

                }

                try {
                    
                    $uiHash.StatusText = "Running audio quality test {0}/{1}" -f $testIteration, $totalIterations
                    $uiHash.ProgressValue = $testIteration / ($totalIterations + 1) * 100
                    Invoke-AudioTest -testIteration $testIteration -totalIterations $totalIterations -ErrorVariable AudioTestError
                    
                    $uiHash.StatusText = "Pausing for {0} seconds..." -f $uiHash.AudioTestDelay
                    Start-Sleep -Seconds $uiHash.AudioTestDelay

                } catch {

                    $uiHash.StatusText = $_.Exception.Message
                    throw

                }

                if ($audioTestError) {

                    $uiHash.StatusText = $audioTestError
                    throw

                }
            }
            
            Invoke-ProcessAudioResults

            $uiHash.ProgressValue = 0
            $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.barTest.Visibility = "Hidden"})
            $uiHash.StatusText = "Click the start button to begin testing"

        }

        if ($uiHash.ActionButtonText -eq "Install") {
            Invoke-NetworkAssessmentToolServicing -ServicingType Install
            Invoke-CheckToolInstall
        }

        if ($uiHash.ActionButtonText -eq "Update") {
            Invoke-NetworkAssessmentToolServicing -ServicingType Uninstall
            Invoke-NetworkAssessmentToolServicing -ServicingType Install
            Invoke-CheckToolInstall
        }

        $variableHash.TestingIdle = $true        
    }

    if ($uiHash.ActionButtonText -eq "Stop") {

        $uiHash.ActionButtonText = "Stopping"
        $uiHash.StopTesting = $true        

    } elseif ($uiHash.ActionButtonText -eq "Running") {

        #do nothing

    } else {

        Invoke-NewRunspace -codeBlock $actionBlock -RunspaceHandleName RspActionButton

    }
    
}

function Invoke-AudioTest {
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)][int]$testIteration,
        [parameter(Mandatory=$true)][int]$totalIterations
    )
    
    #set maximum time limit for the process to run. It should only take about 17s to execute
        
    $stopWatch = [Diagnostics.Stopwatch]::StartNew()
    [int]$timeLimit = 30
    $assessmentInstance = Start-Process -FilePath $variableHash.networkAssessmentToolPath -WorkingDirectory $variableHash.networkAssessmentToolWorkingDirectory -WindowStyle Hidden -PassThru
    
    $uiHash.ActionButtonText = "Stop"

    do {

        Start-Sleep -Milliseconds 500

    } until (($assessmentInstance.HasExited -eq $true) -or ($stopWatch.Elapsed.TotalSeconds -ge $timeLimit))

    $stopWatch.Stop()
    
    #did it complete or did we have to exit?
    if ($assessmentInstance.HasExited) {

        #read the results from the file into memory
        if ($variableHash.Delimiter -ne ",") {
            $results = Import-Csv -Path $variableHash.AudioTestSavePath -Delimiter `t
        } else {
            $results = Import-Csv -Path $variableHash.AudioTestSavePath -Delimiter $variableHash.Delimiter
        }
        
        foreach ($row in $results) {

            $row.CallStartTime = (([datetime]::parse($row.CallStartTime)).ToUniversalTime()).tostring('u')
            $row.PacketLossRate = [single]::parse($row.PacketLossRate)
            $row.RoundTripLatencyInMs = [single]::parse($row.RoundTripLatencyInMs)
            $row.RoundTripLatencyInMs = [int]$row.RoundTripLatencyInMs
            $row.PacketsSent = [int]$row.PacketsSent
            $row.PacketsReceived = [int]$row.PacketsReceived
            $row.AverageJitterInMs = [single]::parse($row.AverageJitterInMs)
            $row.PacketReorderRatio = [single]::parse($row.PacketReorderRatio)

            [array]$variableHash.TestCallData += $row
        }
    } else {
        #kill it
        $assessmentInstance.Kill()
        $uiHash.StatusText = "The test took longer than $timeLimit seconds and was killed."
        Start-Sleep -Seconds 3
    }
}

function Invoke-ConnectivityTest {
    
    $uiHash.StatusText = "Running connectivity test"
    $uiHash.ActionButtonText = "Running"
    $stopWatch = [Diagnostics.Stopwatch]::StartNew()
    [int]$timeLimit = $uiHash.ConnectivityTimeout #we've observed some PC's taking up to 66 seconds to complete this test

    $connTestArgs = "/connectivitycheck /verbose"
    $assessmentInstance = Start-Process -FilePath $variableHash.networkAssessmentToolPath -WorkingDirectory $variableHash.networkAssessmentToolWorkingDirectory -ArgumentList $connTestArgs -WindowStyle Hidden -PassThru -RedirectStandardOutput $variableHash.connectivityCheckResults
    
    do {
        
        $uiHash.ProgressValue = $stopWatch.Elapsed.Seconds / $timeLimit * 100
        Start-Sleep -Milliseconds 100

    } until (($assessmentInstance.HasExited -eq $true) -or ($stopWatch.Elapsed.TotalSeconds -ge $timeLimit))

    $stopWatch.Stop()
    
    #did it complete or did we have to exit?
    if ($assessmentInstance.HasExited) {

        $uiHash.StatusText = "Completed connectivity test"
        $uiHash.ProgressValue = 100
        Start-Sleep -Seconds 2
        #upload results

    } else {

        #kill it
        $assessmentInstance.Kill()
        $uiHash.ActionButtonText = "The test took longer than $timeLimit to complete and was killed."
        $uiHash.StatusText = "The test timed out before it completed."
        $uiHash.ActionButtonText = "Start"
        throw

    }

    #read results from file
    try {
        $connTestResults = Get-Content -Path $variableHash.connectivityCheckResults
    } catch {

        $uiHash.ActionButtonText = "Could not retrieve connectivity check results!"
        $uiHash.ActionButtonFill = "Red"
        throw

    }

    if ($connTestResults.Contains("Verifications completed successfully")) {

        $uiHash.TestConnectivitySource = $uiHash.checkMark
        $uiHash.TestConnectivityDetailSource = $uiHash.checkMark
        $uiHash.ConnectivityDetailText = "No issues detected. The tool was able to reach all transport relays on all required ports."

    } else {

        $uiHash.TestConnectivitySource = $uiHash.errorMark
        $uiHash.TestConnectivityDetailSource = $uiHash.errorMark
        $uiHash.ConnectivityDetailText = "There was a problem connecting to one or more transport or media relays on the required ports. Click the report icon to view the results."

    }

    $uiHash.ProgressValue = 0

}

function Invoke-CheckToolInstall {

    if (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object DisplayName -eq $variableHash.toolName){
        try {
            [xml]$xmlFile = Get-Content -Path $variableHash.networkAssessmentToolConfig
        } catch {
            $uiHash.StatusText = "Error reading config XML file."
            $uiHash.ActionButtonFill = "Red"
            $uiHash.ActionButtonText = "Error"
            throw
        }
        
        $variableHash.Delimiter = ($xmlFile.configuration.AppSettings.add | Where-Object Key -eq "Delimiter").Value
        $variableHash.NumTestsFromXml = ($xmlFile.Configuration.Appsettings.Add | Where-Object Key -eq NumIterations).Value
        $variableHash.AudioTestSavePath = ($xmlFile.Configuration.Appsettings.Add | Where-Object Key -eq ResultsFilePath).value

        if ((Split-Path -Path $variableHash.AudioTestSavePath) -eq "") {

            $variableHash.AudioTestSavePath = $env:LOCALAPPDATA + "\Microsoft Skype for Business Network Assessment Tool\" + $variableHash.AudioTestSavePath

        }

        $variableHash.ConnectionTestSavePath = ($xmlFile.Configuration.Appsettings.Add | Where-Object Key -eq OutputFilePath).value
        if ((Split-Path -Path $variableHash.ConnectionTestSavePath) -eq "") {
            $variableHash.ConnectionTestSavePath = $env:LOCALAPPDATA + "\Microsoft Skype for Business Network Assessment Tool\" + $variableHash.ConnectionTestSavePath
        }

        $installedVersion = Get-Item $variableHash.networkAssessmentToolPath

        if ([System.Version]$installedVersion.VersionInfo.FileVersion -lt [System.Version]$variableHash.ToolVersion) {

            #needs update
            $msg = "The version detected was {0} and needs to be updated to version {1}" -f  $installedVersion.VersionInfo.FileVersion, $variableHash.ToolVersion
            $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtTool.Text = $msg})
            $uiHash.StatusText = "Click the Update button to upgrade the Network Assessment Tool."
            $uiHash.ActionButtonText = "Update"
            $uiHash.ActionButtonFill = "#FF80CC28" #green
            $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgTool.Source = $uiHash.warningMark})
            $variableHash.ToolReady = $false

        } else {

            #good to go; but make sure the config file is modified

            $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtTool.Text = "The version detected on this computer is compatible with this application."})
            $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgTool.Source = $uiHash.checkMark})

            $variableHash.ToolReady = $true

        }
    } else {

        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtTool.Text = "The Network Assessment Tool was not detected. Click the install button to begin the installation."})
        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgTool.Source = $uiHash.warningMark})
        $uiHash.ActionButtonText = "Install"
        $uiHash.ActionButtonFill = "#FF80CC28" #green
        $variableHash.ToolReady = $false

    }

    $variableHash.TestingIdle = $true
}

function Invoke-InternetConnectionTest {

    $variableHash.InternetReady = $false

    try {

        $Socket = New-Object System.Net.Sockets.TCPClient
        $Connection = $Socket.BeginConnect($variableHash.InternetTestIp, 443, $null, $null)
        $Connection.AsyncWaitHandle.WaitOne(2000,$false) | Out-Null

    } catch {

        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtInternet.Text = "There was a problem executing the test for an Internet connection"})
        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgInternet.Source = $uiHash.errorMark})

    }

    if ($Socket.Connected) {

        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtInternet.Text = "Your Internet connection has been successfully verified."})
        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgInternet.Source = $uiHash.checkMark})
        $variableHash.InternetReady = $true

    } else {

        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtInternet.Text = "Unable to verify Internet connectivity to $($variableHash.InternetTestIp). An Internet connection is necessary to start testing."})
        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgInternet.Source = $uiHash.errorMark})

    }

    $Socket.Close | Out-Null

}

function Invoke-ProcessAudioResults {

    #gather test results in memory, write to file, then run results analyzer against it
    $uiHash.ProgressValue = $uiHash.ProgressValue + $uiHash.ProgressValue / 2
    $uiHash.StatusText = "Exporting results from {0} test(s)" -f $variableHash.TestCallData.Count
    Start-Sleep -Seconds 1

    try {
        ($variableHash.TestCallData | ConvertTo-Csv -Delimiter `t -NoTypeInformation).replace('"',"") | Out-File -FilePath $variableHash.resultsAnalyzerTempFile
    } catch {
        $uiHash.StatusText = "There was an error exporting the results: {0}" -f $_.Exception.Message
        Start-Sleep -Seconds 2
        throw
    }
    
    try {
        $uiHash.ProgressValue = 100
        $uiHash.StatusText = "Analyzing audio test results"
        Start-Sleep -Seconds 1
        $argList = '"{0}" {1}' -f $variableHash.resultsAnalyzerTempFile, "`t"
        $networkAssessmentResultsAnalyzer = "${env:ProgramFiles(x86)}\Microsoft Skype for Business Network Assessment Tool\ResultsAnalyzer.exe"
        Start-Process -FilePath $networkAssessmentResultsAnalyzer -WorkingDirectory $variableHash.networkAssessmentToolWorkingDirectory -ArgumentList $argList -PassThru -Wait -WindowStyle Hidden -RedirectStandardOutput $variableHash.resultsAnalyzerTextFile
    } catch {
        $uiHash.StatusText = "There was a problem starting the results analyzer tool."
        throw
    }

    #parse txt file
    $tempResults = Get-Content -Path $variableHash.resultsAnalyzerTextFile

    $uiHash.packetLossRate = ((($tempResults | Where-Object {$_ -like "Packet loss rate:*"})[1] -split ":")[1]).trim()
    $uiHash.PacketLossRateText = ((($tempResults | Where-Object {$_ -like "Packet loss rate:*"})[0] -split ":")[1]).trim()
    if ($uiHash.packetLossRate -eq "PASSED") {
        $uiHash.PacketLossRateSource = $uiHash.checkMark
    } else {
        $uiHash.PacketLossRateSource = $uiHash.errorMark
    }

    $uiHash.rttLatency = ((($tempResults | Where-Object {$_ -like "RTT latency:*"})[1] -split ":")[1]).trim()
    $uiHash.RoundTripTimeText = ((($tempResults | Where-Object {$_ -like "RTT latency:*"})[0] -split ":")[1]).trim()
    if ($uiHash.rttLatency -eq "PASSED") {
        $uiHash.RoundTripTimeSource = $uiHash.checkMark
    } else {
        $uiHash.RoundTripTimeSource = $uiHash.errorMark
    }

    $uiHash.jitter = ((($tempResults | Where-Object {$_ -like "Jitter:*"})[1] -split ":")[1]).trim()
    $uiHash.JitterText = ((($tempResults | Where-Object {$_ -like "Jitter:*"})[0] -split ":")[1]).trim()
    if ($uiHash.jitter -eq "PASSED") {
        $uiHash.JitterSource = $uiHash.checkMark
    } else {
        $uiHash.JitterSource = $uiHash.errorMark
    }

    $uiHash.packetReorderRatio = ((($tempResults | Where-Object {$_ -like "Packet reorder ratio:*"})[1] -split ":")[1]).trim()
    $uiHash.PacketReorderRatioText = ((($tempResults | Where-Object {$_ -like "Packet reorder ratio:*"})[0] -split ":")[1]).trim()
    if ($uiHash.packetReorderRatio -eq "PASSED") {
        $uiHash.PacketReorderRatioSource = $uiHash.checkMark
    } else {
        $uiHash.PacketReorderRatioSource = $uiHash.errorMark
    }

    if ($uiHash.packetLossRate -eq "PASSED" -and $uiHash.rttLatency -eq "PASSED" -and $uiHash.jitter -eq "PASSED" -and $uiHash.packetReorderRatio -eq "PASSED") {
        $uiHash.TestQualitySource = $uiHash.checkMark
        $uiHash.TestQualityDetailSource = $uiHash.checkMark
    } else {
        $uiHash.TestQualitySource = $uiHash.errorMark
        $uiHash.TestQualityDetailSource = $uiHash.errorMark
    }

    $variableHash.TestCallData = $null
    $uiHash.StopTesting = $false

}

function Invoke-CheckCertifiedEndpoint {

    [bool]$isDeviceCertified = $false
    
    if (Test-Path -path $($variableHash.RootPath + "\Deviceparserstandalone.dll")) {

        try {

            Add-Type -Path $($variableHash.RootPath + "\Deviceparserstandalone.dll")
            [array]$usbAudio = Get-WmiObject Win32_PnPEntity | Where-Object {$_.Service -eq 'usbaudio' -and $_.Status -eq "OK"}
            $usbAudio | ForEach-Object {

                if ([SkypeHelpers.CQTools.DeviceParser+Audio]::IsCertified($_.Caption)) {

                    [array]$certifiedDevices += $_.Caption

                }

            }

            if ($certifiedDevices) {

                $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtHeadset.Text = "A Teams/Skype-certified device was detected."})
                $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgHeadset.Source = $uiHash.checkMark})

            } else {

                $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtHeadset.Text = "A Teams/Skype-certified device was not detected. Please click the info icon to learn more."})
                $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgHeadset.Source = $uiHash.warningMark})

            }
            
        } catch {

            $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtHeadset.Text = "There was an error loading the necessary file in memory to determine the headset type."})
            $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgHeadset.Source = $uiHash.warningMark})

        }
    } else {

        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.txtHeadset.Text = "Could not locate the necessary files to determine the headset type."})
        $uiHash.Window.Dispatcher.Invoke([action]{$uiHash.imgHeadset.Source = $uiHash.warningMark})

    }
}

function Invoke-NetworkAssessmentToolServicing {
    [CmdletBinding()]
    param(
        [parameter()][validateset("Install","Uninstall")]$ServicingType
    )
    
    begin{
        #set function constants
        $modulePath = $variableHash.RootPath
        $toolFileName = $modulePath + "\MicrosoftSkypeForBusinessNetworkAssessmentTool.exe"
        $toolInstallArgs = "/install /quiet /norestart /log $env:TEMP\ToolInstall.txt"
        $toolUninstallArgs = "/uninstall /quiet /norestart /log $env:TEMP\ToolUninstall.txt"
    }

    process{
        if ($ServicingType -eq "Uninstall") {
            
            #is it installed already?
            $uiHash.StatusText = "Looking for existing {0} so we can remove it..." -f $variableHash.toolName
            $installedResult = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq $variableHash.toolName -and $_.BundleProviderKey -ne $null}

            if (!$installedResult) {

                $uiHash.StatusText = "{0} was not found on this system." -f $variableHash.toolName
                Start-Sleep -Seconds 1
                return

            } else {
                #yes, remove it first
                $uiHash.StatusText = "Attempting to remove existing {0}..." -f $variableHash.toolName
                Start-Sleep -Seconds 1

                try {

                    $uninstallResult = Start-Process -FilePath $installedResult.BundleCachePath -ArgumentList $toolUninstallArgs -NoNewWindow -PassThru
                    
                    do {

                        $uiHash.StatusText = "Waiting for the tool uninstall to complete..."
                        Start-Sleep -Seconds 1

                    } while ($uninstallResult.HasExited -eq $false)

                } catch {

                    throw

                }

                $installedResult = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object {$_.DisplayName -eq $variableHash.toolName -and $_.BundleProviderKey -ne $null}

                if (!$installedResult) {
                    
                    [xml]$xmlFile = Get-Content -Path $variableHash.networkAssessmentToolConfig
                    ($xmlFile.Configuration.Appsettings.Add | Where-Object Key -eq NumIterations).Value = "1"
                    ($xmlFile.Configuration.Appsettings.Add | Where-Object Key -eq "Delimiter").Value = ","
                    $xmlFile.Save($variableHash.networkAssessmentToolConfig)

                    $uiHash.StatusText = "Successfully uninstalled {0}." -f $variableHash.toolName
                    Start-Sleep -Seconds 1

                } else {

                    $uiHash.StatusText = "There was an error during the removal."
                    $uiHash.ActionButtonFill = "Red"
                    $uiHash.ActionButtonText = "Error"
                    Start-Sleep -Seconds 3
                    throw

                }
            }
        }
        
        if ($ServicingType -eq "Install"){
            try{
                #look for source files first!
                $uiHash.ProgressValue = 30
                Get-Item $toolFileName | Out-Null
                $uiHash.StatusText = "Attempting to install the tool..."
                $installProcess = Start-Process -FilePath $toolFileName -ArgumentList $toolInstallArgs -NoNewWindow -PassThru

                do {
                    $uiHash.ProgressValue = 65
                    $uiHash.StatusText = "Waiting for the tool installation to complete..."
                    Start-Sleep -Seconds 1
                } while ($installProcess.HasExited -eq $false)

            }catch{
                #missing source file .exe
                $uiHash.StatusText = "Could not find installation file!"
                $uiHash.ProgressValue = 0
                $uiHash.ActionButtonFill = "Red"
                $uiHash.ActionButtonText = "Error"
                throw;
            }

            #need to verify it has been installed
            $uiHash.ProgressValue = 80
            $uiHash.StatusText = "Verifying installation..."
            Start-Sleep -Seconds 1
            $installedResult = Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Where-Object DisplayName -eq $variableHash.toolName

            if ($installedResult) {
                $uiHash.ProgressValue = 100
                $uiHash.StatusText = "Successfully installed the tool!"
                Start-Sleep -Seconds 1 
                $uiHash.ToolImageSource = $uiHash.checkMark
            } else {
                $uiHash.StatusText = "Could not verify tool installation!"
                $uiHash.ActionButtonFill = "Red"
                $uiHash.ActionButtonText = "Error"
            }

            $uiHash.ProgressValue = 0
        }
    }

}

function Invoke-NetworkTestingCompanionAutoUpdate {
    [cmdletbinding()]
    param(
        [switch]$CheckForUpdates,
        [switch]$Update
    )

    $codeUpdateWindow = {
        
        $uiHash.Hosts.("RspUpdateWindow") = $Host
        Add-Type -AssemblyName PresentationFramework
        
        $uiContent = Get-Content -Path ($variableHash.rootPath + "\Update.xaml")
        
        [xml]$xAML = $uiContent -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
        $xmlReader = (New-Object System.Xml.XmlNodeReader $xAML)
        $uiHash.UpdateWindow = [Windows.Markup.XamlReader]::Load($xmlReader)

        $xAML.SelectNodes("//*[@Name]") | ForEach-Object {
            $uiHash.Add($_.Name, $uiHash.UpdateWindow.FindName($_.Name))
        }

        $uiHash.UpdateWindow.Add_SourceInitialized(
            {
                $assets = $variableHash.RootPath + "\assets\"
                $uiHash.imgDL.Source = $variableHash.RootPath + "\assets\download.png"
                $uiHash.UpdateWindow.Icon = $assets + "msft_logo.png"
                $uiHash.txtUpdateInstructions.Text = ""

                if ($uiHash.AdminMode) {

                    $uiHash.UpdateInstructionsText = "Click the download icon above to upgrade to version {0}." -f $variableHash.psgVersion

                    $codeUpdateRefresh = {

                        $uiHash.txtUpdateInstructions.Text = $uiHash.UpdateInstructionsText
    
                    }
    
                    #Create timer to handle updating the UI
                    $timer = new-object System.Windows.Threading.DispatcherTimer
                    $timer.Interval = [TimeSpan]"0:0:0:0.30"
                    $timer.Add_Tick($codeUpdateRefresh)
                    $timer.Start()

                } else {

                    $uiHash.imgDL.Cursor = $null
                    $uiHash.btnCloseUpdate.Visibility = "Visible"
                    $uiHash.txtUpdateInstructions.Text = "Please launch this application as an Administrator to update"

                }

            }
        )

        $uiHash.btnCloseUpdate.Add_Click(
            {

                $uiHash.UpdateWindow.Close()

            }
        )

        $uiHash.imgDL.Add_MouseUp(
            {
                
                if ($uiHash.AdminMode) {

                    $codeUpgradeModule = {

                        $uiHash.btnCloseUpdate.IsEnabled = $false
                        $uiHash.UpdateInstructionsText = "Updating the module..."
                    
                        try {
                
                            Update-Module $variableHash.ThisModule.Name -Force -Confirm:$false
                            Invoke-NetworkTestingCompanionCleanup
                
                        } catch {
                
                            $uiHash.UpdateInstructionsText = "Failure executing Update-Module"
                            break;
                
                        }
                        
                        $uiHash.UpdateInstructionsText = "Validating installation..."
                        $availableModules = Get-Module $variableHash.ThisModule.Name -ListAvailable
    
                        if (($availableModules | Sort-Object Version -Descending)[0].Version -eq $variableHash.psgVersion) {
    
                            Invoke-CreateShortcuts
                            $uiHash.UpdateInstructionsText = "Update successful! Please close and re-launch the application."
                            $uiHash.btnCloseUpdate.IsEnabled = $true
    
                            do {
                                Start-Sleep -Seconds 10
                            } until ($infinity)
    
                        } else {
    
                            $uiHash.UpdateInstructionsText = "Update unsuccessful."
                            $uiHash.btnCloseUpdate.IsEnabled = $true
    
                        }

                    }

                    Invoke-NewRunspace -codeBlock $codeUpgradeModule -RunspaceHandleName UpgradeModule

                }

            }
        )

        $uiHash.UpdateWindow.ShowDialog() | Out-Null
        $uiHash.UpdateWindow.Error = $Error

    }

    $uiHash.CheckUpdateEnabled = $false

    if ($CheckForUpdates) {

        $uiHash.SplashStatusText = "Please wait while we check for updates online..."
        [System.Version]$variableHash.psgVersion = (Find-Module $variableHash.ThisModule.Name).Version #time consuming!
        $uiHash.CurrentVersionText = "Current version: {0}" -f $variableHash.ThisModule.Version
        $uiHash.PSGVersionText = "Online version: {0}" -f $variableHash.psgVersion

        if ($variableHash.psgVersion -gt [System.Version]$variableHash.ThisModule.Version) {
            
            Invoke-NewRunspace -codeBlock $codeUpdateWindow -RunspaceHandleName UpdateWindow

            do {

                Start-Sleep -Seconds 1

            } until ($uiHash.Handles.UpdateWindow.IsCompleted)

        } else {

            $uiHash.SplashStatusText = "Up to date"
            Start-Sleep -Seconds 1

        }

    }

    $uiHash.CheckUpdateEnabled = $true

}

function Invoke-NetworkTestingCompanionCleanup {

    #remove older versions N-2
    
    $uiHash.UpdateInstructionsText = "Performing cleanup..."
    Start-Sleep -Seconds 1

    for ($i = 1; $i -le $variableHash.allModules.Count-1; $i++) {

        $uiHash.UpdateInstructionsText = "Removing old module {0}" -f $variableHash.allModules[$i].Version.ToString()
        Start-Sleep -Seconds 1
        
        try {
            
            Uninstall-Module $variableHash.allModules[$i].Name -RequiredVersion $variableHash.allModules[$i].Version -Confirm:$false -Force -ErrorVariable UpdateError
            
        } catch {

            $uiHash.UpdateInstructionsText = "Unable to remove version {0}" -f $variableHash.allModules[$i].Version.ToString()
            Start-Sleep -Seconds 1

        }

        if ($UpdateError) {

            $uiHash.UpdateInstructionsText = "Unable to remove version {0}" -f $variableHash.allModules[$i].Version.ToString()
            Start-Sleep -Seconds 1

        }

    }

}

function Invoke-CalculateEstimatedTime {

    if ($uiHash.NumAudioTests -eq 1) {

        $timeCalc = [math]::Round((40 / 60),1)

    } else {

        $timeCalc = [math]::Round($(([int]$uiHash.NumAudioTests * (40 + [int]$uiHash.AudioTestDelay)) / 60),1) # 40s/test

    }

    $uiHash.txtEstimatedTimeToRun.Text = "Estimated time to complete tests (minutes): " + $timeCalc
    
}

function Get-ToolRegistryEntries {

    $m365RegPath = "HKEY_CURRENT_USER\Software\Microsoft\NetworkTestingCompanion"
    
    $tempKey = [Microsoft.Win32.Registry]::GetValue($m365RegPath,"AutoUpdate","")

    if ($tempKey -eq "False" -or [string]::IsNullOrEmpty($tempKey)) {
        $tempKey = $false
    } else {
        $tempKey = $true
    }
    
    $m365RegKeys = [PSCustomObject]@{
        AutoUpdate = $tempKey
        NumberOfAudioTestIterations = [Microsoft.Win32.Registry]::GetValue($m365RegPath,"NumberOfAudioTestIterations","")
        TestIntervalInSeconds = [Microsoft.Win32.Registry]::GetValue($m365RegPath,"TestIntervalInSeconds","")
        ConnectivityTimeoutInSeconds = [Microsoft.Win32.Registry]::GetValue($m365RegPath,"ConnectivityTimeoutInSeconds","")
    }

    return $m365RegKeys
}

function Set-ToolRegistryEntries {

    $m365RegPath = "HKEY_CURRENT_USER\Software\Microsoft\NetworkTestingCompanion"

        try {

            [Microsoft.Win32.Registry]::SetValue($m365RegPath,"AutoUpdate",[bool]$uiHash.AutoUpdate)
            [Microsoft.Win32.Registry]::SetValue($m365RegPath,"NumberOfAudioTestIterations",$uiHash.cmbNumAudioTests.SelectedItem)
            [Microsoft.Win32.Registry]::SetValue($m365RegPath,"TestIntervalInSeconds",$uiHash.cmbAudioTestDelay.SelectedItem)
            [Microsoft.Win32.Registry]::SetValue($m365RegPath,"ConnectivityTimeoutInSeconds",$uiHash.cmbConnectivityTimeout.SelectedItem)

        } catch {

            throw; #need better handling here

        }
}

function Invoke-ToolCreateShortcuts {
    
    New-Item -ItemType Directory -Path $($env:APPDATA + "\Microsoft\Windows\Start Menu\Programs\Network Testing Companion") -ErrorAction SilentlyContinue | Out-Null
    
    $thisModule = Get-Module NetworkTestingCompanion
    $rootPath = Split-Path -Path $thisModule.Path
    $appLocation = '%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe'
    $appArguments = '-WindowStyle Hidden -Command & {Invoke-NetworkTestingClient}'
    $WshShell = New-Object -ComObject WScript.Shell

    try {

        $Shortcut = $WshShell.CreateShortcut($env:USERPROFILE + "\Desktop\Network Testing Companion.lnk")
        $Shortcut.TargetPath = $appLocation
        $Shortcut.Arguments = $appArguments
        $Shortcut.IconLocation = $rootPath + "\assets\msft_logo.ico"
        $Shortcut.WindowStyle = 7
        $Shortcut.Save()
        Write-Output "Successfully created Desktop shortcut."

    } catch {

        throw
        
    }

    try {

        $StartShortcut = $WshShell.CreateShortcut($env:APPDATA + "\Microsoft\Windows\Start Menu\Programs\Network Testing Companion\Network Testing Companion.lnk")
        $StartShortcut.TargetPath = $appLocation
        $StartShortcut.Arguments = $appArguments
        $StartShortcut.IconLocation = $rootPath + "\assets\msft_logo.ico"
        $StartShortcut.WindowStyle = 7
        $StartShortcut.Save()
        Write-Output "Successfully created Start Menu shortcut."

    } catch {

        throw

    }

}