function New-ARConfigurationFile{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [string]$ConfigFile
    )
    # Configuration Parameters
    $config = @{
        'Domain' = ''
        'ADServer' = ''
        'InitialNotifyDays' = '15'
        'DailyNotifyDays' = '5'
        'SMTPServer' = ''
        'SMTPPort' = '25'
        'SMTPUseSSL' = 'false'
        'MailFrom' = ''
        'MailSubject' = 'Password Expiry Notification'
        'AdminMail' = ''
        'MailUser' = ''
        'MailPassword' = ''
        'ADUser' = ''
        'ADPassword' = ''
    }

    try{
    # Reading existing configuration file in to $config
        $xmlConfigFile = New-Object System.Xml.XmlDocument
        $xmlConfigFile.load($ConfigFile)
        if($xmlConfig = $xmlConfigFile.SelectSingleNode('/Settings')){
            foreach($key in $($config.Keys)){
                if($xmlConfig.$key){
                    $config[$key] = $xmlConfig.$key
                }
            }
        }
    }
    catch{
    #Nothing to do here
    }
    finally{
    # Creating new configuration file from $config
        $xmlConfigFile = New-Object System.Xml.XmlDocument
        $xmlConfigFile.LoadXml('<?xml version="1.0"?><Settings></Settings>')
        foreach($key in $($config.Keys)){
            $msg = 'Please enter ' + $key + '[' + $config[$key] + ']'
            if($key -eq 'ADUser' -OR $key -eq 'MailUser'){
            # Setting the AD user and mail user credential
                $credential = Get-Credential -Message $msg
                if($credential -AND !([string]::IsNullOrWhiteSpace($credential.GetNetworkCredential().Password))){
                    $config[$key] = $credential.UserName
                    if($key -eq 'ADUser'){
                        $config['ADPassword'] = ($credential.Password | ConvertFrom-SecureString)
                        $configKeys = @($key, 'ADPassword')
                    }
                    else{
                        $config['MailPassword'] = ($credential.Password | ConvertFrom-SecureString)
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
                    $config[$key] = $value
                }
                $configKeys = @($key)
            }
            foreach($configKey in $configKeys){            
                $xmlConfigKey = $xmlConfigFile.CreateElement($configKey)
                $xmlConfigValue = $xmlConfigFile.CreateTextNode($config[$configKey])
                $xmlConfigKey.AppendChild($xmlConfigValue) | Out-Null
                $xmlConfigFile.SelectSingleNode('/Settings').AppendChild($xmlConfigKey) | Out-Null
            }
            
        }
        $xmlConfigFile.Save($configFile)
    }
}

$configFile = '.\Settings.xml'
New-ARConfigurationFile -ConfigFile $configFile
