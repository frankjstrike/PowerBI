param (
    [Parameter(Mandatory=$true)]
    [string]$username,
    [Parameter(Mandatory=$true)]
    [string]$pwdTxt
)

Import-Module MicrosoftPowerBIMgmt

# Create credential object
$secureStringPwd = ConvertTo-SecureString $pwdTxt -AsPlainText -Force
$credObject = New-Object System.Management.Automation.PSCredential($username, $secureStringPwd)

#set folder path
$path = "\\Path\To\PowerBI\Workspaces"

# set backup directory
$backup = "\\path\to\backup\directory"

#set logging
$log = $backup +"\_logs\" + (Get-Date -Format "MM-dd-yyyy_HH-mm" ) + ".log" 

Connect-PowerBIServiceAccount -Credential $credObject

# Iterate through directories (workspaces) in $path
Get-ChildItem -Directory -Path $path | foreach {

    $workspaceName = $_.Name
    $workspacePath = $_.FullName

    #Find all PBIX files in each Workspace's root level
    Get-ChildItem -File -Path $_.FullName -Filter "*.pbix" | foreach {
            
        Start-Transcript -path $log -Append

        try {
            Write-Host "*************************Processing $($_.Name)*************************"
                
            #PBIX file info
            $report = $_.BaseName
            $pathToPBIX = $_.FullName

            #Gather IDs from PowerBI
            $workspaceId =  Get-PowerBIWorkspace -Name $workspaceName | Select-Object -ExpandProperty Id

            Write-Host "`n`t> Workspace : $workspaceName"
            Write-Host "`t> Report: $report"
            Write-Host "`t> Workspace ID: $workspaceId"            

            # Confirm that local workspace exists on PowerBI, if not throw exception to skip report
            if ( $workspaceId -eq $null ) {
                Throw "`n> $workspace does not exist on PowerBI side"
            }

            #Get Report and Dataset IDs
            $ErrorActionPreference="SilentlyContinue" #Suppress error if report does not exist on PowerBI side to avoid getting continues errors while waiting for report to be created
            $reportId = Get-PowerBIReport -workspaceId $workspaceId -Name $report | Select-Object -ExpandProperty Id
            $datasetId = Get-PowerBIDataset -WorkspaceId $workspaceId | Where-Object { $_.Name -eq "$report" } | Select-Object -ExpandProperty Id

            # Confirm that local report exists on PowerBI's workspace, If so, back it up
            if ( $reportId -eq $null ) {
                Write-Host "`n> $report does not exists on in the $workspaceName workspace"
            } else {
                #Backup current PBIX file from remote Workspace
                Write-Host "`n> Backing up $report on PowerBI locally to $backup\$WorkSpaceNameDated"
                $WorkSpaceNameDated = $report + "-" + (Get-Date -Format "MM-dd-yyyy_HH-mm" ) + ".pbix"
                $downloadpbixurl = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/reports/$reportId/Export"
                Invoke-PowerBIRestMethod -Url $downloadpbixurl -Method GET -ContentType application/zip -OutFile $backup\$WorkSpaceNameDated

                #Wait for backup to be completed to avoid corruption of backup file
                while ( !(Test-Path $backup\$WorkSpaceNameDated) ) {
                    Write-Host "`n`t> Waiting for $report to be backed up locally"
                    Start-Sleep -Seconds 5
                }
            }

            #Upload PBIX to PowerBI in corresponding workspace
            Write-Host "`n> Uploading $report to $workspaceName"
            New-PowerBIReport -Path "$pathToPBIX" -workspaceId $workspaceId -Name "$report" -ConflictAction CreateOrOverwrite -Timeout 999

            #Wait for report to be created in case it does not exist on PowerBI side
            while ( $reportId -eq $null ) {
                Write-Host "`n> Waiting for $report to be created in $workspaceName"
                Start-Sleep -Seconds 5
                $reportId = Get-PowerBIReport -workspaceId $workspaceId -Name $report | Select-Object -ExpandProperty Id
                $datasetId = Get-PowerBIDataset -WorkspaceId $workspaceId | Where-Object { $_.Name -eq "$report" } | Select-Object -ExpandProperty Id
            }

            #Set error action back to continue
            $ErrorActionPreference="Continue" 

            Write-Host "`n> Gathering Report and Dataset IDs"
            Write-Host "`n`t> Report ID: $reportId"
            Write-Host "`t> Dataset ID: $datasetId"

            #Get Gateway ID
            Write-Host "`n> Gathering GatewayId"
            $get_gateways_url = "https://api.powerbi.com/v1.0/myorg/datasets/$datasetId/Default.DiscoverGateways"
            $gatewayJson = Invoke-PowerBIRestMethod -Url $get_gateways_url -Method GET | ConvertFrom-Json
            $gatewayIds = $gatewayJson.value.id

            # Add Each Gateway ID to an array
            $gatewayIdsArray = @()
            foreach ($id in $gatewayIds) {
                $gatewayIdsArray += $id
            }

            Write-Host "`n`t> Gateway IDs: $gatewayIdsArray"
                
            # Get DatasourceID
            Write-Host "`n> Gathering DatasourceId"
            $get_datasource_url = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId/Default.GetBoundGatewayDatasources"
            $datasourceJson = Invoke-PowerBIRestMethod -Url $get_datasource_url -Method GET | ConvertFrom-Json
            $datasourceIds = $datasourceJson.value.id

            # In case there are multiple datasource ids, create an array to pass it on the call to map the gateway later
            $datasourceIdsArray = @()
            foreach ($id in $datasourceIds) {
                $datasourceIdsArray += $id
            }

            Write-Host "`n`t> Datasource IDs: $datasourceIdsArray"

            # Map Gateway to Datasource
            Write-Host "`n> Mapping Gateway to Datasource"
            $bind_gateway_url = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId/Default.BindToGateway"
                
            # Map the datasources for each gateway
            foreach ($gatewayid in $gatewayIdsArray) {

                $bind_gateway_body = @{
                    "gatewayObjectId" = "$gatewayId"
                    "datasourceObjectIds" = $datasourceIdsArray
                } | ConvertTo-Json

                Invoke-PowerBIRestMethod -Url $bind_gateway_url -Method POST -Body $bind_gateway_body -ContentType "application/json" -TimeoutSec 0
            }

            # Enable Scheduled Refresh
            if (Test-Path -Path $workspacePath\$report.json -PathType Leaf) 
            {
                write-host "`n> Schedule JSON File exists"
                #Enable Schedule and set timezone
                $schedule_refresh_url = "https://api.powerbi.com/v1.0/myorg/groups/$workspaceId/datasets/$datasetId/refreshSchedule"
                $schedulebody = @{
                    "value"= @{
                        "enabled" = "true"
                        "localTimeZoneId" = "Eastern Standard Time"
                        }
                } | ConvertTo-Json
                Invoke-PowerBIRestMethod -Url $schedule_refresh_url -Method PATCH -Body $schedulebody -ContentType application/json
                        
                #Update Schedule Run Times
                Write-Host "`t`n> Updating Schedule refresh"
                $schedule_update_body = Get-Content $workspacePath\$report.json -Raw
                Invoke-PowerBIRestMethod -Url $schedule_refresh_url -Method PATCH -Body $schedule_update_body -ContentType application/json
            }
            else
            {
                write-host "`n> NO JSON FILE FOUND"
            }
            
        
        } catch {
            
            # Provide more information about the failure.
            Write-Host "`n> Error Uploading Dataset. Reason: $($_.Exception.Message)"

        } finally {
            Write-Host "`n***********************************************************************"
            
            #PBIX cleanup
            Remove-Item -Path $pathToPBIX -Force
            Stop-Transcript
        }
    }
}

Disconnect-PowerBIServiceAccount
