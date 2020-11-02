## PasswordExpiryNotification

.SYNOPSIS 
</br>Send password expiry notification email to Active Directory users listed in the input file.

.DESCRIPTION
</br>Script will check for password expiry of users specified in the input file and send password expiry notification email.
</br>All the configuration parameters are stored in the configuration file - Settings.xml. 
</br>Download the folder Send-ARPasswordExpiryNotification and change working directory to the folder
</br>
</br>**Usage**
>***cd Send-ARPasswordExpiryNotification
</br>Create input file Users.csv in this folder
</br>Import-Module ./Send-ARPasswordExpiryNotification.psm1
</br>Send-ARPasswordExpiryNotification -Configure # Generate the configuration file.
</br>Send-ARPasswordExpiryNotification***

.INPUTS
* Base Path: $Path (Default is current working directory)
* Configuration File: $Path/Settings.xml
* Input File: $Path/Users.csv
* Log File: ./Application.log
   
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
</br>**Send-ARPasswordExpiryNotification -Path 'C:\PasswordExpiryNotification'**
