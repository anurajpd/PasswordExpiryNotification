function New-ARConfigurationFile{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [hashtable]$ConfigParam,
        [Parameter(Mandatory=$true)]
        [string]$ConfigFile
    )
    New-Item $ConfigFile -Force -ItemType File | Out-Null
    Set-Content -Path $ConfigFile -Value '<?xml version="1.0"?>'
    Add-Content -Path $ConfigFile -Value '<Settings></Settings>'
    $xmlConfig = New-Object System.Xml.XmlDocument
    $xmlConfig.load($ConfigFile)
    foreach($param in $ConfigParam.Keys){
        $paramNode = $xmlConfig.CreateElement($param)
        $xmlConfig.SelectSingleNode('/Settings').AppendChild($paramNode) | Out-Null
    }
    $xmlConfig.Save($configFile)
}
### Processing configuration file
# Configuration Parameters expected in Settings.xml file.
$configParams = @{
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
$configFile = '.\Settings.xml'
if(!(Test-Path $configFile)){
    New-ARConfigurationFile -ConfigFile $configFile -ConfigParam $configParams
}
