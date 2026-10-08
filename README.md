# Get-HMSCloudWatchErrorLogs

A PowerShell script to search for error patterns in AWS CloudWatch Logs.

The script runs a CloudWatch Logs Insights query for each search pattern and exports the results to an Excel file.

## Configuration

Before running the script, update the configuration variables at the beginning of `Get-HMSCloudWatchErrorLogs.ps1`.

Example:

```powershell
$AwsProfile = 'your-aws-profile'
$AwsRegion = 'ap-northeast-1'
$LogGroup = '/aws/containerinsights/your-eks-name/application'

$ErrorPatterns = @(
    'Exception',
    'OOMKilled',
    'SIGSEGV'
)
```

Set the values for your environment:

- `$AwsProfile`: AWS CLI profile name
- `$AwsRegion`: AWS region
- `$LogGroup`: CloudWatch Log Group to search
- `$ErrorPatterns`: Error strings or patterns to search for

Make sure that `$LogGroup` points to the correct environment, such as DEV or PROD.

## Requirements

Before running the script:

- AWS CLI must be installed.
- The AWS CLI profile must be configured.
- The AWS profile must have permission to read CloudWatch Logs.
- Microsoft Excel must be installed.
- `_template.xlsx` must be placed in the same directory as `Get-HMSCloudWatchErrorLogs.ps1`.

Example:

```text
Get-HMSCloudWatchErrorLogs.ps1
_template.xlsx
```

The script copies `_template.xlsx` and uses the copy to create the output Excel file.

## Run

Open PowerShell and specify the start and end date/time.

Example:

```powershell
.\Get-HMSCloudWatchErrorLogs.ps1 -StartDateTime "2026/10/02 00:00:00" -EndDateTime "2026/10/03 00:00:00"
```

The script searches CloudWatch Logs for each error pattern and creates an Excel file containing the results.
