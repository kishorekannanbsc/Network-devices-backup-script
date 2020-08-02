# Network-devices-backup-script
This vbscript can be mapped to a button inside securecrt and can be used to take backup from multiple devices with multiple backup commands

The script needs 3 things to be executed
  1.hosts file where all the devics IP address are saved
  2.commands file where all the device commands are saved, please add exit at the end of the file
  3.logs folder where the backup files are saved
  
The script will check the logs folder and will add the backup files with device hostname, device IP address and backup time
Incase if the logs folder / hosts file / commands file is not available in the path where the script is run, the script will create the logs folder / hosts file / commands file  and add the backup files in it

Under user editable section user can update the script path and the password for the script to run
