 
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
    $loop = $true

    Start-Process "https://apps.microsoft.com/store/detail/app-installer/9NBLGGH4NNS1?hl=en-gb&gl=gb&rtc=1"

    while ($loop) {
        $continue = Read-Host "
        If you have installed the app installer and are ready to use this script please type Y to continue
        [C] Continue"

        switch($continue){
            C {
                $loop = $false
                testAppInstaller
            }
        }
    }

}#End of installAppInstaller

Function installOffice{

    #folder Paths 
    $a = Test-Path -Path "C:\PS"
    $b = Test-Path -Path "C:\PS\Office32bit.exe"

    #Statement to check Paths and create a folder if it doesn't exist
    if(!$a){
        new-Item -Path "C:\" -Name "PS" -ItemType "directory"
        Write-Output "Path Created"
    }else{
        Write-Output "Path directory already exists"
    }

    #Download App installer
    if(!$b){
	Write-Output "Downloading Office 365 32-bit"
    (New-Object System.Net.WebClient).DownloadFile("https://github.com/agukbiz2988/SetupScript/raw/main/OfficeSetup32bit.exe", "C:\PS\Office32bit.exe")
    }else{
        Write-Output "File Already Exists"
    }

    #Install office
    Start-Process -FilePath "C:\PS\Office32bit.exe" -Wait

}#Install Office 32 Bit Ends

Function downloadSOS{

    #folder Paths 
    $a = Test-Path -Path "C:\PS"
    $b = Test-Path -Path "C:\PS\SplashtopSOS.exe"

    #Statement to check Paths and create a folder if it doesn't exist
    if(!$a){
        new-Item -Path "C:\" -Name "PS" -ItemType "directory"
        Write-Output "Path Created"
    }else{
        Write-Output "Path directory already exists"
    }

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
            
    #installation of programs will begin depending previous choices
    $installProgramLoop = $true

    write-output "`nInstallation of selected Programs will begin`n"
   
    #List of Programs
    $list = @( "google.chrome", "Adobe.Acrobat.Reader.32-bit", "7zip.7zip")
            
    for($i = 0; $i -le $list.length -1 ; $i++){
             winget install $list[$i] --accept-source-agreements -h
            } 
    
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
                installOffice}
            1 { $installProgramLoop = $false
                winget install "9WZDNCRD29V9" --accept-source-agreements -h --force}
            2 { $installProgramLoop = $false
                winget install "TheDocumentFoundation.LibreOffice" --accept-source-agreements -h --force }
            3 { $installProgramLoop = $false
	    	Write-Host "No office aplication will be installed" }
            Default {
                Write-Host "Sorry i didnt understand that"
            }
        }
    
    }#End of While Loop

}#End of Install Programs

function uninstallPrograms{

	$apps =@(
 		"McAfee"
		"McAfee® Personal Security",
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
      		"OneNote",
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
    $pro = @(	"9N7WSZGCK7M5",
		"Amazon.com.Amazon_343d40qqvtj1t",
    		"{4493F3B0-51DA-11EC-8AA4-3863BB3CB5A8}",
    		"9WZDNCRD29V9",
      		"Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe",
     		"OneNoteFreeRetail - en-gb",
     		"OneNoteFreeRetail - en-us",
       		"OneNoteFreeRetail - cs-cz",
	 	"OneNoteFreeRetail - da-dk",
     		"OneNoteFreeRetail - de-de",
     		"OneNoteFreeRetail - fr-fr",
       		"OneNoteFreeRetail - es-es",
	 	"OneNoteFreeRetail - it-it",
		"OneNoteFreeRetail - nl-nl",
		"O365HomePremRetail - en-gb",
  		"O365HomePremRetail - en-us",
    		"O365HomePremRetail - cs-cz",
      		"O365HomePremRetail - da-dk",
		"O365HomePremRetail - nl-nl",
  		"O365HomePremRetail - fr-fr",
    		"O365HomePremRetail - de-de",
      		"O365HomePremRetail - it-it",
		"O365HomePremRetail - es-es",
  		"O365HomePremRetail - nl-nl",
  		"Dell Pair",
    		"DellInc.DellCustomerConnect_htrsf667h5kn2",
      		"Dell Inc.DellDigitalDelivery_htrsf667h5kn2",
		"DellInc.DellUpdate_htrsf667h5kn2",
  		"DellInc.MyDell_htrsf667h5kn2",
    		"AD2F1837.HPEasyClean_v10z8vjag6ke6",
      		"AD2F1837.HPPCHardwareDiagnosticsWindows_v10z8vjag6ke6",
		"AD2F1837.HPPrivacySettings_v10z8vjag6ke6",
  		"AD2F1837.HPQuickDrop_v10z8vjag6ke6",
    		"AD2F1837.HPSystemInformation_v10z8vjag6ke6",
      		"AD2F1837.myHP_v10z8vjag6ke6"
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
	        winget uninstall --name $program --accept-source-agreements --silent
	}
}#end of uninstall wolf security

function setPower{

    #Change power and timer settings
    Powercfg /Change monitor-timeout-ac 5 #screentimout 5 minutes
    Powercfg /Change monitor-timeout-dc 5 #screentimout 5 minutes
    Powercfg /Change standby-timeout-ac 0 #sleep never
    Powercfg /Change standby-timeout-dc 0 #sleep never
    powercfg /hibernate off
    
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
        [E] Exit Other Programs  
        "
        
        switch($programChoice){
            0 { installOffice   									    }
            1 { winget install VideoLAN.VLC --accept-source-agreements -h --force                           }
            2 { winget install TheDocumentFoundation.LibreOffice  --accept-source-agreements -h --force     }
            3 { winget install Adobe.Acrobat.Reader.32-bit --accept-source-agreements -h --force            }
            4 { winget install google.chrome --accept-source-agreements -h --force                          }
            5 { winget install Mozilla.Firefox --accept-source-agreements -h --force                        }
            6 { winget install 7zip.7zip --accept-source-agreements -h --force                              }
            7 { winget install Dropbox.Dropbox --accept-source-agreements -h --force                        }
	    8 { downloadSOS										    }
            E { 
                #Stop the While Loop
                $otherProgramsLoop = $false
            }
        }
    }
}

function welcomelogo {
    Write-Host "
    =====================================================================
    =====================================================================

        ///////// ///////// ///////// ///   /// /////////
       ///       ///          ///    ///   /// ///   ///
      ///////// /////////    ///    ///   /// /////////
           /// ///          ///    ///   /// ///
    ///////// /////////    ///    ///////// ///

        ///////// ///////// /////////   /// ///////// /////////
       ///       ///       ///    ///  /// ///   ///    ///
      ///////// ///       /////////   /// /////////    /// 
           /// ///       ///    ///  /// ///          ///
    ///////// ///////// ///    ///  /// ///          ///  By Andy Gratton

    =====================================================================
    =====================================================================
    "
}#End welcomeLogo

function listFunctions{
    Write-Host"
    ------------------------------------------
    | All Functions Available in this Script |
    ------------------------------------------ 
    | testAppInstaller                       |
    | installAppInstaller                    | 
    | installOffice                          | 
    | installPrograms                        | 
    | setPower                               | 
    | welcomelogo                            | 
    | scriptMenu                             | 
    ------------------------------------------
    "
}#End listFunctions

function scriptMenu{
    
    Write-Host "Welcome to the Setup Script please follow all intructions`n"

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
            * {
   		downloadSOS
                testAppInstaller
                setPower
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


#Commands to Run Script
clear-host
welcomelogo
scriptMenu
[Microsoft.PowerShell.PSConsoleReadLine]::ClearHistory()
