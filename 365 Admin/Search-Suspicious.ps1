function Search-Suspicious {
$known_programs = @("AcrobatNotificationClient",`
            "activation-service",`
            "ai",`
            "AnyDesk",`
            "AppHelperCap",`
            "AppleMobileDeviceService",`
            "ApplicationFrameHost",`
            "armsvc",`
            "audiodg",`
            "AzVpnAppx",`
            "backgroundTaskHost",`
            "BraveCrashHandler",`
            "BraveCrashHandler64",`
            "CalculatorApp",`
            "chrome",`
            "CompPkgSrv",`
            "conhost",`
            "CptService",`
            "csrss",`
            "ctfmon",`
            "dasHost",`
            "DiagsCap",`
            "dllhost",`
            "dwm",`
            "esif_uf",`
            "ETDCtrl",`
            "ETDService",`
            "explorer",`
            "FileCoAuth",`
            "fontdrvhost",`
            "gener8",`
            "Gener8CrashHandler",`
            "Gener8CrashHandler64",`
            "GoogleCrashHandler",`
            "GoogleCrashHandler64",`
            "Honeygain",`
            "HPAudioSwitch",`
            "HPCommRecovery",`
            "HPSystemEventUtilityHost",`
            "HuaweiHiSuiteService64",`
            "Idle",`
            "igfxCUIServiceN",`
            "igfxEMN",`
            "IntelAudioService",`
            "IntelCpHDCPSvc",`
            "jhi_service",`
            "LMS",`
            "LockApp",`
            "lsass",`
            "LTSVC",`
            "LTSvcMon",`
            "LTTray",`
            "m_agent_service",`
            "mDNSResponder",`
            "Memory Compression",`
            "Microsoft.Photos",`
            "msedge",`
            "MsMpEng",`
            "MsSense",`
            "MusNotifyIcon",`
            "NetworkCap",`
            "notepad",`
            "OfficeClickToRun",`
            "OneApp.IGCC.WinService",`
            "OneDrive",`
            "openvpn",`
            "OUTLOOK",`
            "PhoneExperienceHost",`
            "powershell",`
            "powershell_ise",`
            "Registry",`
            "RstMwService",`
            "RtkAudUService64",`
            "RtkBtManServ",`
            "rundll32",`
            "RuntimeBroker",`
            "ScpService",`
            "ScreenConnect.ClientService",`
            "ScreenConnect.WindowsClient",`
            "SearchApp",`
            "SearchFilterHost",`
            "SearchIndexer",`
            "SearchProtocolHost",`
            "SECOCL64",`
            "SECOMN64",`
            "SecurityHealthService",`
            "SecurityHealthSystray",`
            "SenseNdr",`
            "services",`
            "SgrmBroker",`
            "ShellExperienceHost",`
            "sihost",`
            "smartscreen",`
            "smss",`
            "spoolsv",`
            "sqlwriter",`
            "StartMenuExperienceHost",`
            "svchost",`
            "SysInfoCap",`
            "System",`
            "SystemSettings",`
            "taskhostw",`
            "Teams",`
            "TeamViewer_Service",`
            "TextInputHost",`
            "TouchpointAnalyticsClientService",`
            "uhssvc",`
            "unsecapp",`
            "UserAwarenessHelper",`
            "UserAwarenessService",`
            "UserOOBEBroker",`
            "VBoxSDS",`
            "VBoxSVC",`
            "Video.UI",`
            "VirtualBox",`
            "wgsslvpnc",`
            "wgsslvpnsrc",`
            "wininit",`
            "winlogon",`
            "WINWORD",`
            "WMIC",`
            "WmiPrvSE",`
            "WRCoreService.x64",`
            "WRSA",`
            "WRSkyClient.x64",`
            "WRSvcMetrics.x64",`
            "WUDFHost"
)

$netstat = netstat -ano | Select-String "LISTENING", "ESTABLISHED"
foreach ($line in $netstat) {
    $parts = $line -split " "
    $pidd = [string]$parts[-1]
    Write-host  "$pidd" -ForegroundColor Yellow
    
    $exe = (Get-Process -pid $pidd).Name

    if($exe -in $known_programs) 
        {Write-Host "Port: $($parts[1]) is connected to the known program Name: $exe"` -ForegroundColor Green}
    else{Write-Host "Port: $($parts[1]) is connected to the suspicious program $exe" -ForegroundColor Red}

}


}