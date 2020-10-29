<#
.Synopsis
   Sends password expiry mail notification.
.DESCRIPTION
   Sends password expiry mail notification.
.NOTES
   Created by: Anuraj PD
   Modified:  Anuraj PD

   Changelog:

   To Do:
.PARAMETER Message
   Message is the content that you wish to add to the log file. 
.EXAMPLE
  Example here
.LINK
   
#>
function Write-ARLog{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet('Information','Warning','Error')]
        [string]$Level='Information',
        [Parameter(Mandatory=$false)]
        [string]$Path='.\application.log'
    )

    if(!(Test-Path $Path)){
        $logFile = New-Item $Path -Force -ItemType File 
    }

    $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 

    switch ($Level) { 
        'Error' { 
            $LevelText = 'ERROR:' 
            } 
        'Warning' { 
            $LevelText = 'WARNING:' 
            } 
        'Information' {
            $LevelText = 'INFO:' 
            } 
        } 
     
    # Write log entry to $Path 
    "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
}
function New-ARConfigurationFile{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$ConfigFile,
        [Parameter(Mandatory=$true)]
        [hashtable]$Config
    )

    try{
    # Reading existing configuration file in to $config
        Write-ARLog -Message 'Reading existing configuration file'
        $xmlConfigFile = New-Object System.Xml.XmlDocument
        $xmlConfigFile.load($ConfigFile)
        if($xmlConfig = $xmlConfigFile.SelectSingleNode('/Settings')){
            foreach($key in $($Config.Keys)){
                if($xmlConfig.$key){
                    $Config[$key] = $xmlConfig.$key
                }
            }
        }
    }
    catch{
    #Nothing to do here
        Write-ARLog -Message 'Not able to read configuration file.'
    }
    finally{
    # Creating new configuration file from $config
        Write-ARLog -Message 'Creating new configuration file.'
        $xmlConfigFile = New-Object System.Xml.XmlDocument
        $xmlConfigFile.LoadXml('<?xml version="1.0"?><Settings></Settings>')
        foreach($key in $($Config.Keys)){
            $msg = 'Please enter ' + $key + '[' + $Config[$key] + ']'
            if($key -eq 'ADUser' -OR $key -eq 'MailUser'){
            # Setting the AD user and mail user credential
                $credential = Get-Credential -Message $msg
                if($credential -AND !([string]::IsNullOrWhiteSpace($credential.GetNetworkCredential().Password))){
                    $Config[$key] = $credential.UserName
                    if($key -eq 'ADUser'){
                        $Config['ADPassword'] = ($credential.Password | ConvertFrom-SecureString)
                        $configKeys = @($key, 'ADPassword')
                    }
                    else{
                        $Config['MailPassword'] = ($credential.Password | ConvertFrom-SecureString)
                        $configKeys = @($key, 'MailPassword')
                    }
                }
            }
            elseif($key -eq 'ADPassword' -OR $key -eq 'MailPassword'){
            # Skipping the AD user password and mail user password
                continue
            }
            else{
            # Setting all other configuration parameters
                if($value = Read-Host -Prompt $msg){
                    $Config[$key] = $value
                }
                $configKeys = @($key)
            }
            foreach($configKey in $configKeys){            
                $xmlConfigKey = $xmlConfigFile.CreateElement($configKey)
                $xmlConfigValue = $xmlConfigFile.CreateTextNode($Config[$configKey])
                $xmlConfigKey.AppendChild($xmlConfigValue) | Out-Null
                $xmlConfigFile.SelectSingleNode('/Settings').AppendChild($xmlConfigKey) | Out-Null
            }
            
        }
        $xmlConfigFile.Save($configFile)
        Write-ARLog -Message 'Completed the creation of configuration file.'
    }
}
function Send-ARPasswordExpiryNotification{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$false)]
        [string]$Path='./',
        [Parameter(Mandatory=$false)]
        [switch]$Configure=$false
    )
    $ErrorActionPreference = 'Stop'
    $WarningPreference = 'SilentlyContinue'
    try{
        $configFile = $Path + 'Settings.xml'
        $inputFile = $Path + 'Users.csv'
        $logFile = $Path + 'application.log'
        $outputFile = $Path + 'Report.txt'

        # Configuration Parameters
        $config = @{
            'Domain' = 'test.domain.local'
            'ADServer' = 'adserver.domain.local'
            'InitialNotifyDays' = '15'
            'DailyNotifyDays' = '5'
            'SMTPServer' = 'smtp.domain.local'
            'SMTPPort' = '25'
            'SMTPUseSSL' = 'false'
            'MailFrom' = 'mailfrom@domain.local'
            'MailSubject' = 'Password Expiry Notification'
            'AdminMail' = 'adminmail@domain.local'
            'MailUser' = ''
            'MailPassword' = ''
            'ADUser' = ''
            'ADPassword' = ''
        }

        if($Configure){
            Write-ARLog -Message 'Starting the configuration.'
            New-ARConfigurationFile -ConfigFile $configFile -Config $config
            Write-ARLog -Message 'Completed the configuration successfully. Exiting application'
            exit 0
        }
    
        Write-ARLog -Message 'Starting the execution for password expiry mail notification.'
                
        ### Processing the input file
        try{
            Write-ARLog -Message ('Validating the input file' + $inputFile)
            $correctHeaders = @('UserName', 'Mail')
            $actualHeaders = (Get-Content $inputFile -TotalCount 1).Split(",")
            foreach($actualHeader in $actualHeaders){
                if(!($correctHeaders.Contains($actualHeader))){
                    $msg = 'Input file validation error. ' + $actualHeader + ' not a valid header. Exiting Application.'
                    Write-ARLog -Message $msg -Level 'Error'
                    exit 1
                }
            }
            $users = Import-Csv $inputFile
            if(!$users){
                Write-ARLog -Message 'Empty input file.' -Level 'Error'
                exit 1
            }
            foreach($user in $users){
                if(!$user.UserName -OR !$user.Mail){
                    Write-ARLog -Message 'Input file validation error. Exiting application'
                    exit 1
                }
            }
        }
        catch{
            $msg = 'Input file error, please validate the input file. ' + $PSItem.ToString()
            Write-ARLog -Message $msg -Level 'Error'
        }
        
        
        ### Processing configuration file
        try{
            $xmlConfigFile = New-Object System.Xml.XmlDocument
            $xmlConfigFile.load($configFile)
            $xmlConfig = $xmlConfigFile.SelectSingleNode('/Settings')
            Write-ARLog -Message ('Validating the configuration file' + $configFile)
            foreach($key in ($config.Keys)){
                if(!$xmlConfig.$key){
                    $msg = 'Configuration file validation error. Parameter ' + $param + ' is missing. Exiting Application.'
                    Write-ARLog -Message $msg -Level 'Error'
                    exit 1
                }
            }
        }
        catch{
            $msg = 'Configuration file error. Please regenerate the configuration file. ' + $PSItem.ToString()
            Write-ARLog -Message $msg -Level 'Error'
        }
        
        
        $adServer = $xmlConfig.ADServer
        [int]$initialNotifyDays = $xmlConfig.InitialNotifyDays
        [int]$dailyNotifyDays = $xmlConfig.DailyNotifyDays
        $smtpServer = $xmlConfig.SMTPServer
        [int]$smtpPort = $xmlConfig.SMTPPort
        [boolean]$smtpUseSSL = $xmlConfig.SMTPUseSSL
        $mailFrom = $xmlConfig.MailFrom
        $mailSubject = $xmlConfig.MailSubject
        $adminMailTo = $xmlConfig.AdminMail.split(':')
        [string]$mailUser = $xmlConfig.MailUser
        $mailPassword = $xmlConfig.MailPassword | ConvertTo-SecureString
        [string]$adUser = $xmlConfig.ADUser
        $adPassword = $xmlConfig.ADPassword | ConvertTo-SecureString
        
        $adCredential = New-Object System.Management.Automation.PSCredential ($adUser, $adPassword)
        $mailCredential = New-Object System.Management.Automation.PSCredential ($mailUser, $mailPassword)
    
        Write-ARLog -Message ('Getting users from active directory server ' + $adServer)
        $adUsers = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed", "mail", "PasswordLastSet" -Server $adServer -Credential $adCredential
        if(!$adUsers){
            Write-ARLog -Message 'Active Diretory returned empty user list. Exiting Application.' -Level 'Error'
            exit 1
        }
             
        $today = Get-Date
        $notifiedUsers = [System.Collections.ArrayList]@()
        $pwdExpiredUsers = [System.Collections.ArrayList]@()
        $pwdChangedUsers = [System.Collections.ArrayList]@()
        $pendingUsers = [System.Collections.ArrayList]@()
        foreach($adUser in $adUsers){
            foreach($user in $users){           
                if($user.UserName -eq $adUser.SamAccountName){
                    # HTML Template Email Body.
                    $mailBody = @"
                    <div style="font-family:verdana;">
                        <h3>Dear displayName,</h3>
                        <h5>Greetings!!!</h5>
                        <p>
                            Password of you account <strong>accountName</strong> will expire in <strong>noDays</strong> Days. 
                            </br>Please change you password before <strong>expireDate</strong>.
                            </br>Thank you for your understanding.
                            </br>
                            </br>Please contact the Admin Team for any assistance.
                            </br>
                            </br>Best Regards
                            </br>Admin Team
                            </br>aaaaa@mail.com
                            </br>+974 5555 66666
                        </p>
                    </div>
"@
                    $sendMail = $false
                    $passwordExpiryDate = [datetime]::FromFileTime($adUser."msDS-UserPasswordExpiryTimeComputed")
                    $passwordExpiryDays = (New-TimeSpan -Start $today -End $passwordExpiryDate).Days
                    $passwordChangedDays = (New-TimeSpan -Start $adUser.PasswordLastSet -End $today).Days
                    $mailTo = $user.Mail.Split(':')
                    $mailBody = $mailBody.Replace('displayName',$adUser.DisplayName)
                    $mailBody = $mailBody.Replace('accountName',$adUser.SamAccountName)
                    if($passwordExpiryDays -eq $initialNotifyDays){
                        $sendMail = $true
                    }elseif($passwordExpiryDays -le $dailyNotifyDays){
                        $pendingUsers.Add($adUser.SamAccountName)
                        $sendMail = $true
                    }elseif($passwordExpiryDays -le 0){
                        $pwdExpiredUsers.Add($adUser.SamAccountName)
                    }

                    if($passwordChangedDays -lt $initialNotifyDays){
                        $pwdChangedUsers.Add($adUser.SamAccountName)
                    }
                    
                    if($sendMail){
                        $notifiedUsers.Add($adUser.SamAccountName)
                        Write-ARLog -Message ('Sending mail for user ' + $ad.SamAccountName)
                        Send-MailMessage -From $mailFrom -to $mailTo -Subject $mailSubject -Body $mailBody -SmtpServer $smtpServer -Port $smtpPort -Credential $mailCredential
                    }
        
                }
            
            }
        
        }
        # Send Summary Mail to Admin
        $adminMailBody = @"
        <div style="font-family:Arial;">
            <h3>Dear Admin,</h3>
            <h5>Greetings!!!</h5>
            <p>Please find the below summary report.<p>
            <style type="text/css">
            .tg  {border-collapse:collapse;border-color:#aaa;border-spacing:0;}
            .tg td{background-color:#fff;border-color:#aaa;border-style:solid;border-width:1px;color:#333;
            font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 20px;word-break:normal;}
            .tg th{background-color:#f38630;border-color:#aaa;border-style:solid;border-width:1px;color:#fff;
            font-family:Arial, sans-serif;font-size:14px;font-weight:normal;overflow:hidden;padding:10px 20px;word-break:normal;}
            .tg .tg-0lax{text-align:left;vertical-align:top}
            </style>
            <table class="tg">
                <thead>
                <tr>
                    <th class="tg-0lax">Notified Users</th>
                    <th class="tg-0lax">Password Expired Users</th>
                    <th class="tg-0lax">Password Changed Users</th>
                    <th class="tg-0lax">Pending Users</th>
                </tr>
                </thead>
                <tbody>
                <tr>
                    <td class="tg-0lax">notifiedUserCount</td>
                    <td class="tg-0lax">pwdExpiredUserCount</td>
                    <td class="tg-0lax">pwdChangedUserCount</td>
                    <td class="tg-0lax">pendingUserCount</td>
                </tr>
                <tr>
                    <td class="tg-0lax">notifiedUser</td>
                    <td class="tg-0lax">pwdExpiredUser</td>
                    <td class="tg-0lax">pwdChangedUser</td>
                    <td class="tg-0lax">pendingUser</td>
                </tr>
                </tbody>
            </table>
            <p>
                </br>Best Regards
                </br>Password Expiry Notification Application
            </p>
        </div>
"@
        $adminMailBody = $adminMailBody.Replace('notifiedUserCount',$notifiedUsers.Count)
        $adminMailBody = $adminMailBody.Replace('pwdExpiredUserCount',$pwdExpiredUsers.Count)
        $adminMailBody = $adminMailBody.Replace('pwdChangedUserCount',$pwdChangedUsers.Count)
        $adminMailBody = $adminMailBody.Replace('pendingUserCount',$pendingUsers.Count)

        $adminMailBody = $adminMailBody.Replace('notifiedUser',($notifiedUsers -join '</br>'))
        $adminMailBody = $adminMailBody.Replace('pwdExpiredUser',($pwdExpiredUsers -join '</br>'))
        $adminMailBody = $adminMailBody.Replace('pwdChangedUser',($pwdChangedUsers -join '</br>'))
        $adminMailBody = $adminMailBody.Replace('pendingUser',($pendingUsers -join '</br>'))

        Write-ARLog -Message ('Sending summary mail to admin ' + ($adminMailTo -join ','))
        Send-MailMessage -From $mailFrom -to $adminMailTo -Subject $mailSubject -Body $adminMailBody -SmtpServer $smtpServer -Port $smtpPort -Credential $mailCredential

        Write-ARLog -Message 'Successfully completed the execution of the application.'
    }
    catch{
        Write-ARLog -Message $PSItem.ToString() -Level 'Error'
        Write-ARLog -Message 'Unhandled exception occured. Exiting application.'
        exit 1
    }
}

Send-ARPasswordExpiryNotification -Configure












