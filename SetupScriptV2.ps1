function testAppInstaller{
    try{
        Write-Host "`nWinget/ App Installer"
        #1/0
        winget -v

        Write-Host "`nApp Installer is currently installed on this system`n"

    }catch{
        Write-Warning "
        
        App Installer is not installed on this system!!!

        You will need to install the App Installer Application in order to use this Setup Script`n"

        $answer = Read-Host -Prompt "
        Would you like to install the App Installer to use this Script?
        [Y] Yes [N] No
        "
        switch($answer){
            Y {installAppInstaller}
            N { Write-Warning "`nPlease be aware you will not be able to use most of the scripts functionality with App Installer Installed`n" }
        }
    }
}#Test App Installer Ends

function installAppInstaller{

	write-host "Downloading and installing Winget/App Installer"

    try {
        #Install/update winget using built in methods
        #Winget Upgrade Microsoft.AppInstaller
		
		Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile "$env:TEMP\AppInstaller.msixbundle"
		Add-AppxPackage "$env:TEMP\AppInstaller.msixbundle"

        #Winget -v
    }
    catch {
        Write-Warning "Error installing/updating winget using built in methods"		
    #Add-AppxPackage https://tinyurl.com/w1nget
    # createFolderPath
    
    # Invoke-WebRequest -Uri "https://tinyurl.com/w1nget" -OutFile "C:\PS\wingetinstaller.msixbundle"
    # start-process "C:\PS\wingetinstaller.msixbundle"
             
    }
}#End of installAppInstaller

function installOffice{

        #installation of programs will begin depending previous choices
        $installProgramLoop = $true

        #Install Office 365
        while ($installProgramLoop) {
            $officeChoice = Read-Host "
            Please select which office application you would like to install
            Please choose ONE of the following Options Below:
    
            [0] Office 32 bit
            [1] Office 64 bit
            [2] Libre Office
            [3] No Office
            "
    
            switch ($officeChoice) {
                0 { $installProgramLoop = $false
                    installOffice32}
                1 { $installProgramLoop = $false
                    installOffice64}
                2 { $installProgramLoop = $false
                    winget install "TheDocumentFoundation.LibreOffice" --accept-source-agreements -h --force }
                3 { $installProgramLoop = $false
                Write-Host "No office aplication will be installed" }
                Default {
                    Write-Host "Sorry i didnt understand that"
                }
            }
        
        }#End of While Loop
}

function createFolderPath {

        #folder Paths 
        $a = Test-Path -Path "C:\PS"
        
    
        #Statement to check Paths and create a folder if it doesn't exist
        if(!$a){
            new-Item -Path "C:\" -Name "PS" -ItemType "directory"
            Write-Output "Path Created"
        }else{
            Write-Output "Path directory already exists"
        }
    
}

Function removeAllOffice{

    createFolderPath

    $b = Test-Path -Path "C:\PS\setup.exe"
    $c = Test-Path -Path "C:\PS\configurationUninstall.xml"

    #Download Office 365 Uninstaller
    if(!$b ){
        Write-Output "Downloading Office Products Uninstaller"
        (New-Object System.Net.WebClient).DownloadFile("https://github.com/agukbiz2988/SetupScript/raw/main/setup.exe", "C:\PS\setup.exe")
    }else{
        Write-Output "Setup.exe Already Exists"
    }

    if(!$c ){
        Write-Output "Downloading Office Uninstall Configuration"        
        (New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/agukbiz2988/SetupScript/main/configurationUninstall.xml", "C:\PS\configurationUninstall.xml")
    }else{
        Write-Output "configurationUninstall.xml Already Exists"
    }

        Write-Output "Uninstalling All Office Products"
        try{
            C:\PS\setup.exe /configure C:\PS\configurationUninstall.xml
        }Catch{
            "Sorry there was an error"
        }
        Write-Output "Office Uninstalled"

}#Uninstall office ends

Function installOffice32{

    createFolderPath

    $b = Test-Path -Path "C:\PS\setup.exe"
    $c = Test-Path -Path "C:\PS\configuration32Bit.xml"

    #Download Office 365 installer
    if(!$b){
        Write-Output "Downloading Setup.exe"
        (New-Object System.Net.WebClient).DownloadFile("https://github.com/agukbiz2988/SetupScript/raw/main/setup.exe", "C:\PS\setup.exe")
    }else{
        Write-Output "Setup.exe Already Exists"
    }
    
    if(!$c ){
        Write-Output "Downloading Office 32 Bit Configuration"        
        (New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/agukbiz2988/SetupScript/main/configuration32Bit.xml", "C:\PS\configuration32Bit.xml")
    }else{
        Write-Output "Configuration32Bit.xml Already Exists"
    }

    #Install office
    Write-Output "Installing Office 365 Apps For Business"
    try{
        C:\PS\setup.exe /configure C:\PS\configuration32Bit.xml
    }Catch{
        "Sorry there was an error installing Office 365"
    }
    Write-Output "Office 365 Apps For Business Installed"


}#Install Office 32 Bit Ends


Function installOffice64{

    createFolderPath

    $b = Test-Path -Path "C:\PS\setup.exe"
    $c = Test-Path -Path "C:\PS\configuration.xml"

    #Download Office 365 installer
    if(!$b){
        Write-Output "Downloading Setup.exe"
        (New-Object System.Net.WebClient).DownloadFile("https://github.com/agukbiz2988/SetupScript/raw/main/setup.exe", "C:\PS\setup.exe")
    }else{
        Write-Output "Setup.exe Already Exists"
    }
    
    if(!$c ){
        Write-Output "Downloading Office 64 Bit Configuration"        
        (New-Object System.Net.WebClient).DownloadFile("https://raw.githubusercontent.com/agukbiz2988/SetupScript/main/configuration.xml", "C:\PS\configuration.xml")
    }else{
        Write-Output "Configuration.xml Already Exists"
    }

    #Install office
    Write-Output "Installing Office 365 Apps For Business"
    try{
        C:\PS\setup.exe /configure C:\PS\configuration.xml
    }Catch{
        "Sorry there was an error installing Office 365"
    }
    Write-Output "Office 365 Apps For Business Installed"

}#Install Office 64 Bit Ends



Function downloadSOS{

    #folder Paths 
    createFolderPath

    $b = Test-Path -Path "C:\PS\SplashtopSOS.exe"

    #Download App installer
    if(!$b){
	Write-Output "Downloading SplashtopSOS.exe"
    (New-Object System.Net.WebClient).DownloadFile("https://github.com/agukbiz2988/SetupScript/raw/main/SplashtopSOS.exe", "C:\PS\SplashtopSOS.exe")
    }else{
        Write-Output "File Already Exists"
    }

    #
    Copy-Item "C:\PS\SplashtopSOS.exe" -Destination "C:\Users\Public\Desktop"

}


function installPrograms{
            
    write-output "`nInstallation of selected Programs will begin`n"

    #List of Programs
    $list = @( "google.chrome", "Adobe.Acrobat.Reader.64-bit")
            
    for($i = 0; $i -le $list.length -1 ; $i++){
             winget install $list[$i] -h
        } 
    
    installOffice64   
    
}#End of Install Programs

function uninstallPrograms{

	$apps =@(
 		"McAfee® Personal Security",
 		"McAfee",
  		"McAfee LiveSafe",
  		"WebAdvisor by McAfee",
		"Amazon Alexa",
  		"Disney+",
    	"Spotify Music",
    	"Xbox TCUI",
      	"Xbox Console Companion",
		"Xbox Game Bar Plugin",
  		"Xbox Game Bar",
  		"Xbox Identity Provider",
    	"Xbox Game Speech Window",
      	"Xbox",
      	"Microsoft Solitaire Collection",
		"Netflix",
  		"Prime Video for Windows",
    	"Dropbox promotion",
    	"ExpressVPN"
	);    
	
	foreach($program in $apps){
 		Write-Host "Uninstalling $program"
		winget uninstall --name $program --accept-source-agreements --silent
	}
    
    #List of Programs to uninstall if they're on the system
    $pro = @(	"ARP\Machine\X64\McAfee.WPS",
    		"9N7WSZGCK7M5",
		"Amazon.com.Amazon_343d40qqvtj1t",
    	"{4493F3B0-51DA-11EC-8AA4-3863BB3CB5A8}",
  		"Dell Pair",
    	"DellInc.DellCustomerConnect_htrsf667h5kn2",
      	"Dell Inc.DellDigitalDelivery_htrsf667h5kn2",
    	"AD2F1837.HPEasyClean_v10z8vjag6ke6",
      	"AD2F1837.HPPCHardwareDiagnosticsWindows_v10z8vjag6ke6",
		"AD2F1837.HPPrivacySettings_v10z8vjag6ke6",
  		"AD2F1837.HPQuickDrop_v10z8vjag6ke6",
    	"AD2F1837.HPSystemInformation_v10z8vjag6ke6"
	  );
    
    #Foreach loop to uninstall list of programs
    foreach($program in $pro){
        winget uninstall $program --accept-source-agreements --silent
    }

    #uninstall wolf security
    wolfUninstall

}#End of Uninstall Programs

#uninstall wolf security
function wolfUninstall{
	$hpWolf =@(
	        "HP Wolf Security",
	        "HP Wolf Security - Console",
	        "HP Security Update Service"
	);    

	foreach($program in $hpWolf){
 		Write-Host "Uninstalling $program"
	        winget uninstall --name $program --silent
	}
}#end of uninstall wolf security

function setPower{

    #Change power and timer settings
    Powercfg /Change monitor-timeout-ac 5 #screentimout 5 minutes
    Powercfg /Change monitor-timeout-dc 5 #screentimout 5 minutes
    Powercfg /Change standby-timeout-ac 0 #sleep never
    Powercfg /Change standby-timeout-dc 0 #sleep never
    powercfg /hibernate off

    #Enable Password Prompt After Wake-Up:
    powercfg /SETDCVALUEINDEX SCHEME_CURRENT SUB_NONE CONSOLELOCK 1
    powercfg /SETACVALUEINDEX SCHEME_CURRENT SUB_NONE CONSOLELOCK 1
    powercfg -SetActive SCHEME_CURRENT

    #Sets Time Zone
    Set-TimeZone -Id 'GMT Standard Time'

}#End of SetPower

function otherPrograms{
    
    $otherProgramsLoop = $true
    while ($otherProgramsLoop) {
        $programChoice = Read-Host "`nPlease Select one of the following Programs you would like to install
        
        [0] Office 365
        [1] VLC Media Player
        [2] Libre Office
        [3] Adobe Reader DC
        [4] Google Chrome
        [5] Mozilla Firefox
        [6] 7zip
        [7] Dropbox
	[8] Splashtop SOS
    	[9] Splashtop Streamer
        [10] PowerToys (Preview)
        [11] WindHawk

        [E] Exit Other Programs  
        "
        
        switch($programChoice){
            0 { installOffice   									    }
            1 { winget install VideoLAN.VLC --accept-source-agreements -h --force                           }
            2 { winget install TheDocumentFoundation.LibreOffice  --accept-source-agreements -h --force     }
            3 { winget install Adobe.Acrobat.Reader.64-bit --accept-source-agreements -h --force            }
            4 { winget install google.chrome --accept-source-agreements -h --force                          }
            5 { winget install Mozilla.Firefox --accept-source-agreements -h --force                        }
            6 { winget install 7zip.7zip --accept-source-agreements -h --force                              }
            7 { winget install Dropbox.Dropbox --accept-source-agreements -h --force                        }
	    8 { downloadSOS										    }
     	    9 { winget install Splashtop.SplashtopStreamer -h --force                                       }
            10 { winget install Microsoft.PowerToys -h --force                                              }
            11 { winget RamenSoftware.Windhawk -h --force                                                   }
            E { 
                #Stop the While Loop
                $otherProgramsLoop = $false
            }
        }
    }
}

function Set-PasswordPolicy{
    param(
        [Int]$policy,
        [Int]$num
    )

    switch($policy){
        1{ Net Accounts /MINPWLEN:$num         }
        2{ Net Accounts /MINPWAGE:$num         }
        3{ Net Accounts /MAXPWAGE:$num         }
        4{ Net Accounts /UNIQUEPW:$num         }
        5{ Net accounts /lockoutthreshold:$num }
        6{ Net Accounts /lockoutduration:$num  }
        7{ Net Accounts /lockoutwindow:$num    }
    }
}

function Set-PasswordComplexity{
    param(
        [Int] $num
    )

    secedit.exe /export /cfg C:\secconfig.cfg
    Start-Sleep 1

    switch($num){
        0{
            #Disable Complexity
            (Get-Content -path C:\secconfig.cfg -Raw) -replace 'PasswordComplexity = 1', 'PasswordComplexity = 0' | Out-File -FilePath C:\secconfig.cfg
            Start-Sleep 1
        }
        1{
            #Enable Complexity
            (Get-Content -path C:\secconfig.cfg -Raw) -replace 'PasswordComplexity = 0', 'PasswordComplexity = 1' | Out-File -FilePath C:\secconfig.cfg
            Start-Sleep 1
        }
        default{Write-Host $incorrectNum}
    }
    secedit.exe /configure /db $env:windir\securitynew.sdb /cfg C:\secconfig.cfg /areas SECURITYPOLICY
}

function Set-AdministratorLockout{
    param(
        [Int] $num
    )

    secedit.exe /export /cfg C:\secconfig.cfg
    Start-Sleep 1

    switch($num){
        0{
            #Disable Admin Lockout
            (Get-Content -path C:\secconfig.cfg -Raw) -replace 'AllowAdministratorLockout = 1', 'AllowAdministratorLockout = 0' | Out-File -FilePath C:\secconfig.cfg
            Start-Sleep 1
        }
        1{
            #Enable Admin Lockout
            (Get-Content -path C:\secconfig.cfg -Raw) -replace 'AllowAdministratorLockout = 0', 'AllowAdministratorLockout = 1' | Out-File -FilePath C:\secconfig.cfg
            Start-Sleep 1
        }
        default{Write-Host $incorrectNum}
    }
    secedit.exe /configure /db $env:windir\securitynew.sdb /cfg C:\secconfig.cfg /areas SECURITYPOLICY
}

function Set-DefaultPasswordPolicy{

    #Enable/Disable Password Complexity
    Write-Host "Turning Complexity Off " -ForegroundColor green
    Set-PasswordComplexity 0
    Write-Host "Complextity Turned off`n" -ForegroundColor green
    Start-Sleep 1

    #Enable/Disable Administrator Account Lockout
    Write-Host "Disabling Admin Lockout" -ForegroundColor green
    Set-AdministratorLockout 0
    Write-Host "Completed Disabling Admin Lockout`n" -ForegroundColor green
    Start-Sleep 1

    #Change Minimum Password Length
    Write-Host "Changing Minimum Password Length Back to Default" -ForegroundColor green
    Set-PasswordPolicy -policy 1 -num 0
    Start-Sleep 1

    #Change Minimum Password Age
    Write-Host "Changing Minimum Password Age Back to Default" -ForegroundColor green
    Set-PasswordPolicy -policy 2 -num 0
    Start-Sleep 1

    #Change Maximum Password Age
    Write-Host "Changing Maximum Password Age Back to Default" -ForegroundColor green
    Set-PasswordPolicy -policy 3 -num 42
    Start-Sleep 1

    #Change Password History Size
    Write-Host "Changing Password History Size Back to Default" -ForegroundColor green
    Set-PasswordPolicy -policy 4 -num 0
    Start-Sleep 1

    #Change Lockout Threshold
    Write-Host "Changing Lockout Threshold Back to Default" -ForegroundColor green
    Set-PasswordPolicy -policy 5 -num 10
    Start-Sleep 1

    #Change Lockout Duration Time
    Write-Host "Changing Lockout Duration Time Back to Default" -ForegroundColor green
    Set-PasswordPolicy -policy 7 -num 10
    Start-Sleep 1

    #Change Lockout Window Time
    Write-Host "Changing Lockout Window Time Back to Default" -ForegroundColor green
    Set-PasswordPolicy -policy 6 -num 10
    Start-Sleep 1

    net accounts
}

function Set-RecommendedPasswordPolicy{

    #Enable/Disable Password Complexity
    Write-Host "Turning Complexity On " -ForegroundColor green
    Set-PasswordComplexity 1
    Write-Host "Complextity Turned On`n" -ForegroundColor green
    Start-Sleep 1

    #Enable/Disable Administrator Account Lockout
    Write-Host "Enabling Admin Lockout" -ForegroundColor green
    Set-AdministratorLockout 1
    Write-Host "Completed Enabling Admin Lockout`n" -ForegroundColor green
    Start-Sleep 1

    #Change Minimum Password Length
    Write-Host "Changing Minimum Password Length to Recommended Setting" -ForegroundColor green
    Set-PasswordPolicy -policy 1 -num 12
    Start-Sleep 1

    #Change Minimum Password Age
    Write-Host "Changing Minimum Password Age to Recommended Setting" -ForegroundColor green
    Set-PasswordPolicy -policy 2 -num 0
    Start-Sleep 1

    #Change Maximum Password Age
    Write-Host "Changing Maximum Password Age to Recommended Setting" -ForegroundColor green
    Set-PasswordPolicy -policy 3 -num 182
    Start-Sleep 1

    #Change Password History Size
    Write-Host "Changing Password History Size to Recommended Setting" -ForegroundColor green
    Set-PasswordPolicy -policy 4 -num 24
    Start-Sleep 1

    #Change Lockout Threshold
    Write-Host "Changing Lockout Threshold to Recommended Setting" -ForegroundColor green
    Set-PasswordPolicy -policy 5 -num 10
    Start-Sleep 1

    #Change Lockout Window Time
    Write-Host "Changing Lockout Window Time to Recommended Setting" -ForegroundColor green
    Set-PasswordPolicy -policy 6 -num 30
    Start-Sleep 1

    #Change Lockout Duration Time
    Write-Host "Changing Lockout Duration Time to Recommended Setting" -ForegroundColor green
    Set-PasswordPolicy -policy 7 -num 30
    Start-Sleep 1

    net accounts
}


function welcomelogo {
    
    Write-Host "

    =============================================================================================
    
                          ____    __              ____        _      __ 
                         / __/__ / /___ _____    / __/_______(_)__  / /_
                        _\ \/ -_) __/ // / _ \  _\ \/ __/ __/ / _ \/ __/
                       /___/\__/\__/\_,_/ .__/ /___/\__/_/ /_/ .__/\__/ 
                                       /_/                  /_/         

    				   ** Made by Andy Gratton **

    ==============================================================================================

"
}#End welcomeLogo

welcomelogo

function scriptMenu{
    
    Write-Host "Welcome to the Setup Script please follow all instructions`n"

    $loop = $true

    while ($loop) {

        $choice = Read-Host "
        Please Choose one of the following Options Below:

        [0] Check App installer is Installed
        [1] Install App Installer
        [2] Set Power Settings
        [3] Uninstall All Unnessesary Programs
        [4] Install All Standard Programs
        [5] Other Programs
		[6] Install Splashtop
        [7] Microsoft Office Removal
		[8] Set Recommended Password Policy
        [*] Run all Above Options
        [E] End Script
	[W] Uninstall Wolf Security
        "
        #Switch for Choice Selected
        switch($choice){
            0 { testAppInstaller    }
            1 { installAppInstaller }
            2 { setPower            }
            3 { uninstallPrograms   }
            4 { installPrograms     }
            5 { otherPrograms       }
	        6 { downloadSOS 	    }
            7 { removeAllOffice     }
			8 { Set-RecommendedPasswordPolicy }
            * {
   		        downloadSOS
                installAppInstaller
		        winget list Microsoft.AppInstaller
                setPower
				Set-RecommendedPasswordPolicy
                uninstallPrograms
                installPrograms
            }
            E { 
                #Stop the While Loop
                $loop = $false

                #Remove File Path Created
                $a = Test-Path -Path "C:\PS"
                if($a){
                    Remove-Item -Path "c:\PS" -Recurse
                }

                Write-Host "`nThank you for Using This Setup Script.`n"

		#Clear History
		[Microsoft.PowerShell.PSConsoleReadLine]::ClearHistory()
            }
	    W { wolfUninstall }
        }
    }
}#End sciptMenu


#Commands that run at the start of the script Script
clear-host
welcomelogo
scriptMenu
[Microsoft.PowerShell.PSConsoleReadLine]::ClearHistory()





