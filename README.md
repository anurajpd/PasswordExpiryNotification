## PasswordExpiryNotification
</br>
.SYNOPSIS 
</br>Script to send automatic password expiry mail notification.

.DESCRIPTION
</br>Script looks in the base path for the configuration file, input file and log file.
</br>Script will check for password expiry of users specified in the input file and send 
</br>password expiry notification email.

.INPUTS
* Base Path: $Path or current working directory.
* Base Path:/Settings.xml - Configuration File
* Base Path:/Users.csv - Input File
* Base Path:/Application.log - Log File
   
.NOTES
* Created by: Anuraj PD
* Modified:  Anuraj PD
* Changelog:
* To Do:
    * Log file is hardcoded to current working directory. Change it to the base path.
    * Enable debug mode of logging.
    * Enable log rotation.

.PARAMETER Configure
</br>Create or reconfigure the application configuration file.

.PARAMETER Path
</br>Base path where the application looks for configuration file, input file and log file.

.OUTPUTS
</br>Sends notification mails.

.EXAMPLE
</br>To create the configuration file or reconfigure. The configuration file will be generated at 
the current working directory. 
</br>**Send-ARPasswordExpiryNotification -Configure**

.EXAMPLE
<br>To create the configuration file or reconfigure. The configuration file will be generated at 
the path provided in $Path parameter.
</br>**Send-ARPasswordExpiryNotification -Configure -Path 'C:\PasswordExpiryNotification'***

.EXAMPLE
</br>To send the notification mail. Script will look for configuration file, input file and log file
at the current workingdirectory.
</br>**Send-ARPasswordExpiryNotification**

.EXAMPLE
</br>To send the notification mail. Script will look for configuration file, input file
at the path provided in $Path parameter.
**Send-ARPasswordExpiryNotification -Path 'C:\PasswordExpiryNotification'**
