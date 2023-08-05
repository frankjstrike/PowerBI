# upload_dataset.ps1

This PowerShell script automates the upload, backup and scheduled refresh of PowerBI reports.

## Requirements

- PowerShell
- MicrosoftPowerBIMgmt PowerShell module

## Usage

The script accepts two mandatory parameters:
- username
- pwdTxt (plaintext password)

Here's an example of how to run the script:

```powershell
.\upload_dataset.ps1 -username "example@domain.com" -pwdTxt "yourpassword"
```

## Script Flow

1. The script starts by importing the MicrosoftPowerBIMgmt PowerShell module.
2. A credential object is created using the provided username and password.
3. The script then connects to the PowerBI Service account using the credential object.
4. The script iterates through all directories (considered as workspaces) in a specified path.
5. For each workspace, it finds all PBIX files (PowerBI reports).
6. For each PBIX file, it backs up the current report from PowerBI, uploads the local version of the report to PowerBI, sets up the gateway and data source, and sets up the scheduled refresh if a corresponding JSON file is found.
7. The script then removes the local PBIX file (clean-up step).
8. Finally, the script disconnects from the PowerBI service.


## Logging

The script logs all actions and outputs them to a .log file. This file is located in the specified backup directory and is timestamped for ease of review.

## Notes

* The script handles exceptions by writing an error message to the host.
* The script expects the workspace in PowerBI to have the same name as the directory in the local filesystem.
* Please ensure that the user account has necessary permissions on PowerBI service.




