# Get-HMSCloudWatchErrorLogs

A PowerShell script to search for error patterns in AWS CloudWatch Logs.

The script runs a CloudWatch Logs Insights query for each search pattern and exports the results to an Excel file.

## Configuration

Before running the script, update the configuration variables at the beginning of `Get-HMSCloudWatchErrorLogs.ps1`.

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

## Run

Open PowerShell and run:

```powershell
.\Get-HMSCloudWatchErrorLogs.ps1
```

Enter the start and end date/time when prompted.

The script searches CloudWatch Logs for each error pattern and creates an Excel file containing the results.

## Requirements

- AWS CLI is installed.
- The AWS CLI profile is configured.
- The AWS profile has permission to read CloudWatch Logs.
- `_template.xlsx` is in the same folder as the PowerShell script.
