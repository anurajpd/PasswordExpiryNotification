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