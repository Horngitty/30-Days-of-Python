#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
<#
    Pester tests for Export-UserOneDrive.ps1.

    No Active Directory, Microsoft Graph or network access is needed: the AD and
    Graph cmdlets are replaced with stubs + Pester mocks, and a fake OneDrive
    tree is served from the Invoke-MgGraphRequest mock. Files are written to
    Pester's $TestDrive instead of C:\Temp.

    Run:  Invoke-Pester ./PowerShell/tests -Output Detailed
#>

BeforeAll {
    $script:ScriptPath = Join-Path $PSScriptRoot '..' 'Export-UserOneDrive.ps1'

    # Stubs so the script (and Pester's Mock) can resolve these commands on
    # machines without RSAT / Microsoft.Graph installed.
    if (-not (Get-Command Get-ADUser -ErrorAction SilentlyContinue)) {
        function global:Get-ADUser { param($Identity, $Properties) }
    }
    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        function global:Connect-MgGraph { param($TenantId, $ClientId, $CertificateThumbprint, $Scopes, [switch]$NoWelcome) }
        function global:Disconnect-MgGraph { }
        function global:Invoke-MgGraphRequest { param($Method, $Uri, $OutputFilePath) }
    }

    # Fake OneDrive:
    #   /report.docx            (10 bytes)
    #   /bad:name?.txt          (3 bytes, invalid Windows chars)
    #   /Notebook               (OneNote package - can't be downloaded)
    #   /Projects/              (two pages of children via @odata.nextLink)
    #       plan.xlsx           (20 bytes)
    #       broken.pdf          (download fails)
    $global:OdTestFakeDrive = @{
        'root'    = @(
            @{ value = @(
                    @{ id = 'f1'; name = 'report.docx'; size = 10; file = @{} }
                    @{ id = 'f2'; name = 'bad:name?.txt'; size = 3; file = @{} }
                    @{ id = 'p1'; name = 'Notebook'; size = 0; package = @{ type = 'oneNote' } }
                    @{ id = 'd1'; name = 'Projects'; size = 0; folder = @{ childCount = 2 } }
                ) }
        )
        'd1'      = @(
            @{ value = @(@{ id = 'f3'; name = 'plan.xlsx'; size = 20; file = @{} })
               '@odata.nextLink' = 'v1.0/drives/drive-123/items/d1/children?page=2' }
        )
        'd1-page2' = @(
            @{ value = @(@{ id = 'f4'; name = 'broken.pdf'; size = 5; file = @{} }) }
        )
    }
    $global:OdTestFileSizes = @{ f1 = 10; f2 = 3; f3 = 20; f4 = 5 }

    function script:Invoke-Export {
        param([hashtable]$Params = @{})
        & $script:ScriptPath -DestinationRoot $TestDrive @Params
    }
}

AfterAll {
    Remove-Variable -Scope Global -Name OdTest* -ErrorAction SilentlyContinue
}

Describe 'Export-UserOneDrive.ps1' {

    BeforeEach {
        Get-ChildItem $TestDrive -Force | Remove-Item -Recurse -Force

        Mock Write-Host { }
        Mock Start-Sleep { }
        Mock Import-Module { } -ParameterFilter { $Name -in 'ActiveDirectory', 'Microsoft.Graph.Authentication' }

        Mock Read-Host { 'jsmith' } -ParameterFilter { $Prompt -like 'Enter the AD username*' }
        Mock Read-Host { 'Y' }      -ParameterFilter { $Prompt -like 'Download this user*' }

        Mock Get-ADUser {
            [pscustomobject]@{
                SamAccountName    = 'jsmith'
                UserPrincipalName = 'jsmith@contoso.com'
                DisplayName       = 'John Smith'
                mail              = 'jsmith@contoso.com'
                Enabled           = $true
            }
        }

        Mock Connect-MgGraph { }
        Mock Disconnect-MgGraph { }

        Mock Invoke-MgGraphRequest {
            switch -Regex ($Uri) {
                '/users/.+/drive\?' {
                    return @{ id = 'drive-123'; webUrl = 'https://contoso-my.sharepoint.com/personal/jsmith'; quota = @{ used = 35 } }
                }
                '/items/d1/children\?page=2' { return $global:OdTestFakeDrive['d1-page2'][0] }
                '/items/(\w+)/children'      { return $global:OdTestFakeDrive[$Matches[1]][0] }
                '/items/f4/content' { throw 'The remote server returned an error: (403) Forbidden.' }
                '/items/(\w+)/content' {
                    [IO.File]::WriteAllBytes($OutputFilePath, [byte[]]::new($global:OdTestFileSizes[$Matches[1]]))
                    return
                }
                default { throw "Unexpected Graph URI in test: $Uri" }
            }
        }
    }

    Context 'Happy path' {

        It 'looks up the user in AD by the username entered' {
            Invoke-Export
            Should -Invoke Get-ADUser -Times 1 -Exactly -ParameterFilter { $Identity -eq 'jsmith' }
        }

        It 'requests the OneDrive of the UPN returned by AD' {
            Invoke-Export
            Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Uri -like 'v1.0/users/jsmith@contoso.com/drive*' }
        }

        It 'downloads files into DestinationRoot/username, preserving folders' {
            Invoke-Export
            $root = Join-Path $TestDrive 'jsmith'
            (Get-Item (Join-Path $root 'report.docx')).Length              | Should -Be 10
            (Get-Item (Join-Path $root 'Projects' 'plan.xlsx')).Length     | Should -Be 20
        }

        It 'follows @odata.nextLink paging' {
            Invoke-Export
            Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Uri -like '*items/d1/children?page=2' } -Times 1 -Exactly
        }

        It 'replaces characters that are invalid in Windows file names' {
            Invoke-Export
            Join-Path $TestDrive 'jsmith' 'bad_name_.txt' | Should -Exist
        }

        It 'uses interactive delegated sign-in with the right scopes by default' {
            Invoke-Export
            Should -Invoke Connect-MgGraph -Times 1 -Exactly -ParameterFilter {
                ($Scopes -contains 'Files.Read.All') -and ($Scopes -contains 'User.Read.All') -and -not $CertificateThumbprint
            }
        }

        It 'always disconnects from Graph' {
            Invoke-Export
            Should -Invoke Disconnect-MgGraph -Times 1 -Exactly
        }

        It 'writes a log file' {
            Invoke-Export
            Get-ChildItem $TestDrive -Filter 'jsmith_*.log' | Should -HaveCount 1
        }
    }

    Context 'Failures and skipped items' {

        It 'continues after a failed download and records it in the failures CSV' {
            Invoke-Export
            Join-Path $TestDrive 'jsmith' 'Projects' 'broken.pdf' | Should -Not -Exist

            $csv = Get-ChildItem $TestDrive -Filter 'jsmith_*_failures.csv'
            $csv | Should -HaveCount 1
            $rows = Import-Csv $csv.FullName
            $rows.Path | Should -Contain (Join-Path $TestDrive 'jsmith' 'Projects' 'broken.pdf')
        }

        It 'skips OneNote/package items and lists them as failures' {
            Invoke-Export
            $rows = Import-Csv (Get-ChildItem $TestDrive -Filter 'jsmith_*_failures.csv').FullName
            ($rows | Where-Object Path -like '*Notebook').Error | Should -BeLike '*package*oneNote*'
            Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Uri -like '*items/p1/*' } -Times 0 -Exactly
        }

        It 'retries on HTTP 429 throttling and then succeeds' {
            $global:OdTestCalls = 0
            Mock Invoke-MgGraphRequest {
                $global:OdTestCalls++
                if ($global:OdTestCalls -eq 1) {
                    $resp = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]429)
                    throw [Microsoft.PowerShell.Commands.HttpResponseException]::new('Too Many Requests', $resp)
                }
                return @{ id = 'drive-123'; webUrl = 'x'; quota = @{ used = 0 } }
            } -ParameterFilter { $Uri -like 'v1.0/users/*' }

            Invoke-Export
            $global:OdTestCalls | Should -Be 2
            Should -Invoke Start-Sleep -Times 1 -Exactly
        }

        It 'stops cleanly if the OneDrive cannot be accessed' {
            Mock Invoke-MgGraphRequest { throw 'Access denied' } -ParameterFilter { $Uri -like 'v1.0/users/*' }

            { Invoke-Export } | Should -Not -Throw
            Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Uri -like '*children*' } -Times 0 -Exactly
            Should -Invoke Disconnect-MgGraph -Times 1 -Exactly
        }
    }

    Context 'Resume' {

        It 'skips files that already exist locally with the same size' {
            $root = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'jsmith') -Force
            [IO.File]::WriteAllBytes((Join-Path $root 'report.docx'), [byte[]]::new(10))

            Invoke-Export
            Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Uri -like '*items/f1/content' } -Times 0 -Exactly
            Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Uri -like '*items/f2/content' } -Times 1 -Exactly
        }

        It 're-downloads a file whose local size differs (partial download)' {
            $root = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'jsmith') -Force
            [IO.File]::WriteAllBytes((Join-Path $root 'report.docx'), [byte[]]::new(4))

            Invoke-Export
            Should -Invoke Invoke-MgGraphRequest -ParameterFilter { $Uri -like '*items/f1/content' } -Times 1 -Exactly
            (Get-Item (Join-Path $root 'report.docx')).Length | Should -Be 10
        }
    }

    Context 'User input and AD lookup' {

        It 'throws a clear error when the user is not in AD' {
            Mock Get-ADUser { throw 'Cannot find an object with identity' }
            { Invoke-Export } | Should -Throw "*'jsmith' was not found in Active Directory*"
            Should -Invoke Connect-MgGraph -Times 0 -Exactly
        }

        It 'does nothing when the confirmation is declined' {
            Mock Read-Host { 'n' } -ParameterFilter { $Prompt -like 'Download this user*' }
            Invoke-Export
            Should -Invoke Connect-MgGraph -Times 0 -Exactly
            Join-Path $TestDrive 'jsmith' | Should -Not -Exist
        }

        It 're-prompts when the username is blank' {
            $global:OdTestAnswers = [System.Collections.Generic.Queue[string]]::new([string[]]@('', '   ', 'jsmith'))
            Mock Read-Host { $global:OdTestAnswers.Dequeue() } -ParameterFilter { $Prompt -like 'Enter the AD username*' }

            Invoke-Export
            Should -Invoke Read-Host -ParameterFilter { $Prompt -like 'Enter the AD username*' } -Times 3 -Exactly
        }
    }

    Context 'App-only authentication' {

        It 'uses certificate auth when TenantId, ClientId and thumbprint are supplied' {
            Invoke-Export -Params @{ TenantId = 'contoso.onmicrosoft.com'; ClientId = 'app-id'; CertificateThumbprint = 'ABC123' }
            Should -Invoke Connect-MgGraph -Times 1 -Exactly -ParameterFilter {
                $TenantId -eq 'contoso.onmicrosoft.com' -and $ClientId -eq 'app-id' -and $CertificateThumbprint -eq 'ABC123' -and -not $Scopes
            }
        }
    }
}
