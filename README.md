# Get-CloudWatchErrorLogs

A PowerShell script to search for error patterns in AWS CloudWatch Logs.

The script runs a CloudWatch Logs Insights query for each search pattern and exports the results to an Excel file.

## Configuration

Before running the script, update the configuration variables at the beginning of `Get-CloudWatchErrorLogs.ps1`.

Example:

```powershell
$AwsProfile = 'your-aws-profile'
$AwsRegion = 'ap-northeast-1'
$LogGroup = '/aws/containerinsights/your-eks-name/application'
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
- AWS CLI authentication must be completed, and a valid access token must be available.
- The AWS profile must have permission to read CloudWatch Logs.
- Microsoft Excel must be installed.
- `_template.xlsx` must be placed in the same directory as `Get-CloudWatchErrorLogs.ps1`.

### AWS SSO authentication

If you use AWS IAM Identity Center (SSO), configure your AWS CLI profile with:

```powershell
aws configure sso --no-browser
```

When prompted for the **SSO Start URL**, open the AWS access portal and click **Access keys** for the AWS account and role you want to use. You can find the SSO configuration information there.

With `--no-browser`, the AWS CLI displays a URL for authentication. Open the URL in a browser and follow the instructions to complete authentication.

After the profile is configured, sign in with:

```powershell
aws sso login --profile <your-aws-profile> --no-browser
```

Use the same profile name for `$AwsProfile` in `Get-CloudWatchErrorLogs.ps1`.

### Excel template

`_template.xlsx` must be in the same directory as the PowerShell script.

Example:

```text
Get-CloudWatchErrorLogs.ps1
_template.xlsx
```

The script copies `_template.xlsx` and uses the copy to create the output Excel file.

## Run

Open PowerShell and specify the start and end date/time.

Example:

```powershell
.\Get-CloudWatchErrorLogs.ps1 -StartDateTime "2026/10/02 00:00:00" -EndDateTime "2026/10/03 00:00:00"
```

The script searches CloudWatch Logs for each error pattern and creates an Excel file containing the results.

![Excel File](images/results.png)

