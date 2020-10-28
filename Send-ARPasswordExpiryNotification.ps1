<#
.Synopsis
   Sends password expiry mail notification.
.DESCRIPTION
   Sends password expiry mail notification.
.NOTES
   Created by: 
   Modified:  

   Changelog:

   To Do:
.PARAMETER Message
   Message is the content that you wish to add to the log file. 
.EXAMPLE
  Example here
.LINK
   
#>

function Update-ARCredential{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$false)]
        [Alias('ConfigFilePath')]
        [string]$Path='C:\PasswordExpiryNotification\Settings.xml',
        [Parameter(Mandatory=$false)]
        [switch]$ADUser,
        [Parameter(Mandatory=$false)]
        [switch]$MailUser
    )
    
    if(!(Test-Path $Path)){
        Write-Error "Not able to find the Configuration File $Path"
    }

    $xmlConfig = New-Object System.Xml.XmlDocument
    $xmlConfig.load($Path)
    $config = $xmlConfig.SelectSingleNode('/Settings')

    if($ADUser -AND $MailUser){
    # Both ADUser and MailUser switch is provided, set both AD User and Mail User Credential
        $adCredential = Get-Credential -Message 'Credential for AD User'
        $mailCredential = Get-Credential -Message 'Credential for Mail User'
        if($adCredential -AND $mailCredential){
            if([string]::IsNullOrWhiteSpace($adCredential.GetNetworkCredential().Password)){
                Write-Error 'AD User Password Empty'
            }
            elseif([string]::IsNullOrWhiteSpace($mailCredential.GetNetworkCredential().Password)){
                Write-Error 'Mail User Password Empty'
            }
            else{
                $config.ADUser = $adCredential.UserName
                $config.ADPassword = ($adCredential.Password | ConvertFrom-SecureString)
                $config.MailUser = $mailCredential.UserName
                $config.MailPassword = ($mailCredential.Password | ConvertFrom-SecureString)
                $xmlConfig.Save($Path)
            }
        }
        else{
            Write-Verbose 'User cancelled the credential prompt for AD User or Mail User'
        }
    }
    elseif($ADUser){
    # Only ADUser switch is provided, set only AD User Credential
        $credential = Get-Credential -Message 'Credential for AD User'
        if($credential){
            if([string]::IsNullOrWhiteSpace($credential.GetNetworkCredential().Password)){
                Write-Error 'Empty Password'
            }
            else{
                $config.ADUser = $credential.UserName
                $config.ADPassword = ($credential.Password | ConvertFrom-SecureString)
                $xmlConfig.Save($Path)
            }
        }
        else{
            Write-Verbose 'User cancelled the credential prompt for AD User'
        }
    }
    elseif($MailUser){
    # Only MailUser switch is provided, set only Mail User Credential
        $credential = Get-Credential -Message 'Credential for Mail User'
        if($credential){
            if([string]::IsNullOrWhiteSpace($credential.GetNetworkCredential().Password)){
                Write-Error 'Empty Password'
            }
            else{
                $config.MailUser = $credential.UserName
                $config.MailPassword = ($credential.Password | ConvertFrom-SecureString)
                $xmlConfig.Save($Path)
            }
        }
        else{
            Write-Verbose 'User cancelled the credential prompt for Mail User'
        }
    }
    else{
    # No switch is provided, set the same credential for AD User and Mail User
        $credential = Get-Credential -Message 'Credential for AD User and Mail User'
        if($credential){
            if([string]::IsNullOrWhiteSpace($credential.GetNetworkCredential().Password)){
                Write-Error 'Empty Password'
            }
            else{
                $config.ADUser = $credential.UserName
                $config.ADPassword = ($credential.Password | ConvertFrom-SecureString)
                $config.MailUser = $credential.UserName
                $config.MailPassword = ($credential.Password | ConvertFrom-SecureString)
                $xmlConfig.Save($Path)
            }
        }
        else{
            Write-Verbose 'User cancelled the credential prompt for AD User and Mail User'
        }
        
    }
}

function Send-ARPasswordExpiryNotification{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$false)]
        [Alias('ConfigFilePath')]
        [string]$Path='C:\PasswordExpiryNotification\',
        [Parameter(Mandatory=$false)]
        [switch]$ADUser=$false,
        [Parameter(Mandatory=$false)]
        [switch]$MailUser=$false,
        [Parameter(Mandatory=$false)]
        [switch]$Credential=$false

    )

    $configFile = $Path + 'Settings.xml'
    $inputFile = $Path + 'Users.csv'
    $logFile = $Path + 'log.log'
    $outputFile = $Path + 'Report.txt'

    if($Credential){
        Update-ARCredential -Path $configFile -ADUser $ADUser -MailUser $MailUser
        Write-Output 'Successfully configured the credentials'
        exit 0
    }

    $configParams = @(
        'Domain',
        'ADServer',
        'InitialNotifyDays',
        'DailyNotifyDays',
        'SMTPServer',
        'SMTPPort',
        'SMTPUseSSL',
        'MailFrom',
        'MailSubject',
        'MailUser',
        'MailPassword',
        'ADUser',
        'ADPassword'
        )

    if(!(Test-Path $configFile)){
        Write-Error "Not able to find the Configuration File $configFile"
    }
    if(!(Test-Path $inputFile)){
        Write-Error "Not able to find the Input File $inputFile"
    }
    
    $correctHeaders = @('UserName', 'Mail')
    $actualHeaders = (Get-Content $inputFile -TotalCount 1).Split(",")
    foreach($actualHeader in $actualHeaders){
        if(!($correctHeaders.Contains($actualHeader))){
            Write-Error($actualHeader + ' not a valid header')
        }
    }

    $users = Import-Csv $inputFile
    
    $today = Get-Date

    $xmlConfig = New-Object System.Xml.XmlDocument
    $xmlConfig.load($configFile)
    $config = $xmlConfig.SelectSingleNode('/Settings')
    if(!$config){
        Write-Error 'Configuration File Error. Please verify configuration file'
    }
    
    foreach($param in $configParams){
        if(!$config.$param){
            Write-Error($param + ' configuration Missing from the configuration file')
        }
    }
    
    $adServer = $config.ADServer
    [int]$initialNotifyDays = $config.InitialNotifyDays
    [int]$dailyNotifyDays = $config.DailyNotifyDays
    $smtpServer = $config.SMTPServer
    [int]$smtpPort = $config.SMTPPort
    $smtpUseSSL = $config.SMTPUseSSL
    $mailFrom = $config.MailFrom
    $mailSubject = $config.MailSubject
    [string]$mailUser = $config.MailUser
    $mailPassword = $config.MailPassword | ConvertTo-SecureString
    [string]$adUser = $config.ADUser
    $adPassword = $config.ADPassword | ConvertTo-SecureString
    
    $adCredential = New-Object System.Management.Automation.PSCredential ($adUser, $adPassword)
    $mailCredential = New-Object System.Management.Automation.PSCredential ($mailUser, $mailPassword)

    $adUsers = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed", "mail" -Server $adServer -Credential $adCredential
  
    foreach($ad in $adUsers){
        foreach($user in $users){           
            if($user.UserName -eq $ad.SamAccountName){
                $passwordExpiryDate = [datetime]::FromFileTime($ad."msDS-UserPasswordExpiryTimeComputed")
                $passwordExpiryDays = (New-TimeSpan -Start $today -End $passwordExpiryDate).Days
                $mailTo = $user.Mail.Split(':')
                if($passwordExpiryDays -eq $initialNotifyDays){
                # Sent Initial Notification
                    $mailBody = 'Your Password will Expire in ' + $passwordExpiryDays + ' Days. Please change your Password.'
                    Send-MailMessage -From $mailFrom -to $mailTo -Subject $mailSubject -Body $mailBody -SmtpServer $smtpServer -Port $smtpPort -Credential $mailCredential
                    Write-Output('Sending Mail for User ' + $adUser.SamAccountName)
                }elseif($passwordExpiryDays -le $dailyNotifyDays){
                # Send Daily Notification
                    $mailBody = 'Your Password will Expire in ' + $passwordExpiryDays + ' Days. Please change your Password.'
                    Send-MailMessage -From $mailFrom -to $mailTo -Subject $mailSubject -Body $mailBody -SmtpServer $smtpServer -Port $smtpPort -Credential $mailCredential
                    Write-Output('Sending Mail for User ' + $adUser.SamAccountName)
                }elseif($passwordExpiryDays -le 0){
                # Pass0word Expired  
                    $mailBody = 'Your Password is Expired. Please change your Password.'
                    Send-MailMessage -From $mailFrom -to $mailTo -Subject $mailSubject -Body $mailBody -SmtpServer $smtpServer -Port $smtpPort -Credential $mailCredential
                    Write-Output 'Password Expired'
                }
    
            }
        
        }
    
    }

}

#Send-ARPasswordExpiryNotification












