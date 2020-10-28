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
        Write-ARLog -Message 'Unable to find the configuration file $Path' -Level 'Error'
        Write-ARLog -Message 'Exiting the application' -Level 'Error'
        exit 1
    }

    Write-ARLog -Message 'Loading configuration file $Path'
    $xmlConfig = New-Object System.Xml.XmlDocument
    $xmlConfig.load($Path)
    $config = $xmlConfig.SelectSingleNode('/Settings')

    if($ADUser -AND $MailUser){
    # Both ADUser and MailUser switch is provided, set both AD User and Mail User Credential
        Write-ARLog 'Updating both AD User and Mail User credential'
        $adCredential = Get-Credential -Message 'Credential for AD User'
        $mailCredential = Get-Credential -Message 'Credential for Mail User'
        if($adCredential -AND $mailCredential){
            if([string]::IsNullOrWhiteSpace($adCredential.GetNetworkCredential().Password)){
                Write-ARLog -Message 'Empty password provided for AD User. Exiting the application.' -Level 'Error'
                exit 1
            }
            elseif([string]::IsNullOrWhiteSpace($mailCredential.GetNetworkCredential().Password)){
                Write-ARLog -Message 'Empty password provided for Mail User. Exiting the application.' -Level 'Error'
                exit 1
            }
            else{
                $config.ADUser = $adCredential.UserName
                $config.ADPassword = ($adCredential.Password | ConvertFrom-SecureString)
                $config.MailUser = $mailCredential.UserName
                $config.MailPassword = ($mailCredential.Password | ConvertFrom-SecureString)
                $xmlConfig.Save($Path)
                Write-ARLog -Message 'Updated both AD User and Mail User credential'
            }
        }
        else{
            Write-ARLog -Message 'User cancelled the password prompt. Exiting the application.' -Level 'Error'
            exit 1
        }
    }
    elseif($ADUser){
    # Only ADUser switch is provided, set only AD User Credential
        Write-ARLog 'Updating AD User credential'
        $credential = Get-Credential -Message 'Credential for AD User'
        if($credential){
            if([string]::IsNullOrWhiteSpace($credential.GetNetworkCredential().Password)){
                Write-ARLog -Message 'Empty password provided for AD User. Exiting the application.' -Level 'Error'
                exit 1
            }
            else{
                $config.ADUser = $credential.UserName
                $config.ADPassword = ($credential.Password | ConvertFrom-SecureString)
                $xmlConfig.Save($Path)
                Write-ARLog -Message 'Updated AD User credential'
            }
        }
        else{
            Write-ARLog -Message 'User cancelled the password prompt. Exiting the application.' -Level 'Error'
            exit 1
        }
    }
    elseif($MailUser){
    # Only MailUser switch is provided, set only Mail User Credential
        Write-ARLog 'Updating Mail User credential'
        $credential = Get-Credential -Message 'Credential for Mail User'
        if($credential){
            if([string]::IsNullOrWhiteSpace($credential.GetNetworkCredential().Password)){
                Write-ARLog -Message 'Empty password provided for Mail User. Exiting the application.' -Level 'Error'
                exit 1
            }
            else{
                $config.MailUser = $credential.UserName
                $config.MailPassword = ($credential.Password | ConvertFrom-SecureString)
                $xmlConfig.Save($Path)
                Write-ARLog -Message 'Updated Mail User credential'
            }
        }
        else{
            Write-ARLog -Message 'User cancelled the password prompt. Exiting the application.' -Level 'Error'
            exit 1
        }
    }
    else{
    # No switch is provided, set the same credential for AD User and Mail User
        Write-ARLog 'Updating same credential for both AD User and Mail User'
        $credential = Get-Credential -Message 'Credential for AD User and Mail User'
        if($credential){
            if([string]::IsNullOrWhiteSpace($credential.GetNetworkCredential().Password)){
                Write-ARLog -Message 'Empty password provided. Exiting the application.' -Level 'Error'
                exit 1
            }
            else{
                $config.ADUser = $credential.UserName
                $config.ADPassword = ($credential.Password | ConvertFrom-SecureString)
                $config.MailUser = $credential.UserName
                $config.MailPassword = ($credential.Password | ConvertFrom-SecureString)
                $xmlConfig.Save($Path)
                Write-ARLog -Message 'Updated same credential for both AD User and Mail User'
            }
        }
        else{
            Write-ARLog -Message 'User cancelled the password prompt. Exiting the application.' -Level 'Error'
            exit 1
        }
        
    }
}

function Send-ARPasswordExpiryNotification{
    [CmdletBinding(DefaultParameterSetName = 'Notif')]
    Param
    (
        [Parameter(ParameterSetName = 'Notify', Mandatory=$false)]
        [Parameter(ParameterSetName = 'Credential', Mandatory=$false)]
        [Alias('ConfigFilePath')]
        [string]$Path='C:\PasswordExpiryNotification\',
        [Parameter(ParameterSetName = 'Credential', Mandatory=$false)]
        [switch]$ADUser=$false,
        [Parameter(ParameterSetName = 'Credential', Mandatory=$false)]
        [switch]$MailUser=$false,
        [Parameter(ParameterSetName = 'Credential', Mandatory=$true)]
        [switch]$Credential=$false
    )
    try{
        if($Credential){
            Write-ARLog -Message 'Starting the execution for credential update'
            Update-ARCredential -Path $configFile -ADUser:$ADUser -MailUser:$MailUser
            exit 0
        }
    
        Write-ARLog -Message 'Starting the execution for password expiry mail notification.'
    
        $configFile = $Path + 'Settings.xml'
        $inputFile = $Path + 'Users.csv'
        $logFile = $Path + 'application.log'
        $outputFile = $Path + 'Report.txt'
            
        ### Processing the input file
        Write-ARLog -Message 'Reading input file ' + $inputFile
        if(!(Test-Path $inputFile)){
            $msg = 'Not able to find the input file' + $inputFile + '. Exiting Application.'
            Write-ARLog -Message $msg -Level 'Error'
        }
        Write-ARLog -Message 'Validating the input file' + $inputFile
        $correctHeaders = @('UserName', 'Mail')
        $actualHeaders = (Get-Content $inputFile -TotalCount 1).Split(",")
        foreach($actualHeader in $actualHeaders){
            if(!($correctHeaders.Contains($actualHeader))){
                $msg = 'Input file validation error. ' + $actualHeader + ' not a valid header. Exiting Application.'
                Write-ARLog -Message $msg -Level 'Error'
            }
        }
        $users = Import-Csv $inputFile
        if(!$users){
            Write-ARLog -Message 'Empty input file.' -Level 'Error'
            exit 1
        } 
        
        ### Processing configuration file
        # Configuration Parameters expected in Settings.xml file.
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
            'AdminMail',
            'MailUser',
            'MailPassword',
            'ADUser',
            'ADPassword'
            )
        Write-ARLog -Message 'Reading configuration file ' + $configFile
        if(!(Test-Path $configFile)){
            $msg = 'Not able to find the configuration file' + $configFile +'. Exiting Application.'
            Write-ARLog -Message $msg -Level 'Error'
        }
        $xmlConfig = New-Object System.Xml.XmlDocument
        $xmlConfig.load($configFile)
        $config = $xmlConfig.SelectSingleNode('/Settings')
        if(!$config){
            Write-ARLog -Message 'Configuration file validation error. Exiting Application.' -Level 'Error'
            exit 1
        }
        Write-ARLog -Message 'Validating the configuration file' + $configFile
        foreach($param in $configParams){
            if(!$config.$param){
                $msg = 'Configuration file validation error. Parameter ' + $param + ' is missing. Exiting Application.'
                Write-ARLog -Message $msg -Level 'Error'
                exit 1
            }
        }
        
        $adServer = $config.ADServer
        [int]$initialNotifyDays = $config.InitialNotifyDays
        [int]$dailyNotifyDays = $config.DailyNotifyDays
        $smtpServer = $config.SMTPServer
        [int]$smtpPort = $config.SMTPPort
        [boolean]$smtpUseSSL = $config.SMTPUseSSL
        $mailFrom = $config.MailFrom
        $mailSubject = $config.MailSubject
        $adminMailTo = $config.AdminMail.split(':')
        [string]$mailUser = $config.MailUser
        $mailPassword = $config.MailPassword | ConvertTo-SecureString
        [string]$adUser = $config.ADUser
        $adPassword = $config.ADPassword | ConvertTo-SecureString
        
        $adCredential = New-Object System.Management.Automation.PSCredential ($adUser, $adPassword)
        $mailCredential = New-Object System.Management.Automation.PSCredential ($mailUser, $mailPassword)
    
        Write-ARLog -Message 'Getting users from active directory server ' + $adServer
        $adUsers = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed", "mail", "PasswordLastSet" -Server $adServer -Credential $adCredential
        if(!$adUsers){
            Write-ARLog -Message 'Active Diretory returned empty user list. Exiting Application.' -Level 'Error'
            exit 1
        }
             
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

        $today = Get-Date
        $notifiedUsers = [System.Collections.ArrayList]@()
        $pwdExpiredUsers = [System.Collections.ArrayList]@()
        $pwdChangedUsers = [System.Collections.ArrayList]@()
        $pendingUsers = [System.Collections.ArrayList]@()
        foreach($ad in $adUsers){
            foreach($user in $users){           
                if($user.UserName -eq $ad.SamAccountName){
                    $sendMail = $false
                    $passwordExpiryDate = [datetime]::FromFileTime($ad."msDS-UserPasswordExpiryTimeComputed")
                    $passwordExpiryDays = (New-TimeSpan -Start $today -End $passwordExpiryDate).Days
                    $passwordChangedDays = (New-TimeSpan -Start $ad.PasswordLastSet -End $today).Days
                    $mailTo = $user.Mail.Split(':')
                    $mailBody = $mailBody.Replace('displayName',$ad.DisplayName)
                    $mailBody = $mailBody.Replace('accountName',$ad.SamAccountName)
                    if($passwordExpiryDays -eq $initialNotifyDays){
                        $sendMail = $true
                    }elseif($passwordExpiryDays -le $dailyNotifyDays){
                        $pendingUsers.Add($ad.SamAccountName)
                        $sendMail = $true
                    }elseif($passwordExpiryDays -le 0){
                        $pwdExpiredUsers.Add($ad.SamAccountName)
                    }

                    if($passwordChangedDays -lt $initialNotifyDays){
                        $pwdChangedUsers.Add($ad.SamAccountName)
                    }
                    
                    if($sendMail){
                        $notifiedUsers.Add($ad.SamAccountName)
                        Write-ARLog -Message 'Sending mail for user ' + $ad.SamAccountName
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

        Write-ARLog -Message 'Sending summary mail to admin ' + ($adminMailTo -join ',')
        Send-MailMessage -From $mailFrom -to $adminMailTo -Subject $mailSubject -Body $adminMailBody -SmtpServer $smtpServer -Port $smtpPort -Credential $mailCredential

        Write-ARLog -Message 'Successfully completed the execution of the application.'
    }
    catch{
        Write-ARLog -Message $PSItem.ToString() -Level 'Error'
        Write-ARLog -Message 'Unhandled exception occured. Exiting application.'
        exit 1
    }
}

Send-ARPasswordExpiryNotification -Credential -ADUser -MailUser












