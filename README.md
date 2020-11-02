# PasswordExpiryNotification


<#
.SYNOPSIS
Script to send automatic password expiry mail notification.

.DESCRIPTION
Script looks in the base path for the configuration file, input file and log file.
Script will check for password expiry of users specified in the input file and send 
password expiry notification email.

.INPUTS
Base Path: $Path or current working directory.
Base Path:/Settings.xml - Configuration File
Base Path:/Users.csv - Input File
Base Path:/Application.log - Log File
   
.NOTES
Created by: Anuraj PD
Modified:  Anuraj PD
Changelog:
To Do:
    Log file is hardcoded to current working directory. Change it to the base path.
    Enable debug mode of logging.
    Enable log rotation.

.PARAMETER Configure
Create or reconfigure the application configuration file.

.PARAMETER Path
Base path where the application looks for configuration file, input file and log file.

.OUTPUTS
Sends notification mails.

.EXAMPLE
To create the configuration file or reconfigure. The configuration file will be generated at 
the current working directory. 
Send-ARPasswordExpiryNotification -Configure

.EXAMPLE
To create the configuration file or reconfigure. The configuration file will be generated at 
the path provided in $Path parameter.
Send-ARPasswordExpiryNotification -Configure -Path 'C:\PasswordExpiryNotification'

.EXAMPLE
To send the notification mail. Script will look for configuration file, input file and log file
at the current workingdirectory.
Send_ARPasswordExpiryNotification

.EXAMPLE
To send the notification mail. Script will look for configuration file, input file
at the path provided in $Path parameter.
Send-ARPasswordExpiryNotification -Path 'C:\PasswordExpiryNotification'
#>
