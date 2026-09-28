# Testing the AD / OneDrive scripts

The scripts in this folder need Active Directory and Microsoft 365. Testing them
happens at three levels, cheapest first.

## 1. Unit tests with Pester mocks (no AD or tenant needed)

`tests/*.Tests.ps1` use [Pester 5](https://pester.dev) to replace `Get-ADUser`,
`Connect-MgGraph`, `Invoke-MgGraphRequest` and `Read-Host` with mocks. They serve
a fake OneDrive tree and write files to Pester's temporary `$TestDrive`.
The tests cover paging, invalid file names, failed downloads, OneNote packages,
429 retries, resuming, user-not-found, declining the confirmation, and app-only auth.

```powershell
Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser -Force -SkipPublisherCheck
Invoke-Pester ./PowerShell/tests -Output Detailed
```

These run on Windows, macOS or Linux (PowerShell 7+). They also run in GitHub
Actions (`.github/workflows/powershell-tests.yml`) on every push or PR that
touches `PowerShell/`, together with a PSScriptAnalyzer lint step.

**Adding tests for a new AD script:**
1. If the RSAT or Graph module isn't installed, define a stub function
   (e.g. `function global:Get-ADGroupMember { param($Identity) }`) so `Mock` can find it.
2. `Mock` every cmdlet that touches AD, Graph, the network or the console.
3. Run the script with `& $scriptPath -SomeOutputPath $TestDrive`.
4. Assert with `Should -Invoke` (what was called) and `Should -Exist`/`-Be` (what was written).

## 2. Integration tests in a free sandbox

Mocks only prove that the script's logic is right. They don't prove that the
real cmdlets and permissions behave the way the mocks assume. For that, use a
throwaway environment:

- **Microsoft 365 tenant:** the Microsoft 365 Developer Program sandbox
  (if you qualify), or a trial tenant. Create a test user, upload some files to
  their OneDrive (including nested folders, a OneNote notebook and names with
  `#`, `%` and unicode), then register an app with `Files.Read.All` and
  `User.Read.All` (application) and a certificate.
- **Active Directory:** a Windows Server evaluation VM (180-day, free) promoted
  to a domain controller (`Install-WindowsFeature AD-Domain-Services`,
  `Install-ADDSForest -DomainName lab.local`). Create a user whose UPN matches the
  tenant test user, or use Entra Connect to sync them. Microsoft's
  "AD DS lab" guides and AutomatedLab (`Install-Module AutomatedLab`) can script
  the whole setup.

Then run the real script against the test user and compare the results:

```powershell
.\Export-UserOneDrive.ps1 -DestinationRoot C:\Temp\OneDriveArchiveTest
# Compare the local file count with what Graph reports:
(Get-ChildItem C:\Temp\OneDriveArchiveTest\<user> -Recurse -File).Count
```

## 3. First production run

Run it for one low-risk user and check the log and `_failures.csv` before using
it on anyone else.
