<# 
.SYNOPSIS
    Script for generating all groups used for the ProjectBoards
	
	Execute Script `Powershell.exe -ExecutionPolicy Bypass -File .\AzureAD-Create-ProjectBoards-Groups.ps1`
.DESCRIPTION
    Script for generating all groups used for the ProjectBoards
	
	Execute Script `Powershell.exe -ExecutionPolicy Bypass -File .\AzureAD-Create-ProjectBoards-Groups.ps1`
.NOTES
    Author: Timo Kluser - timo.kluser@edu.gbssg.ch
.LINK
    https://docs.microsoft.com/en-us/graph/overview
    https://docs.microsoft.com/en-us/powershell/module/azuread/?view=azureadps-2.0-preview
#>

################################### Start ###################################
Write-Host '-- Azure AD Groups Creator --'
Write-Host 

################################### Import Modules ###################################
try {
    Import-Module -Name AzureADPreview -ErrorAction Stop
    Import-Module -Name MSAL.PS -ErrorAction Stop
}
catch [System.IO.FileNotFoundException] {
    Write-Host 'Installing AzureADPreview & MSAL.PS' -ForegroundColor Cyan
    Install-Module -Name AzureADPreview -Force
    Import-Module -Name AzureADPreview -ErrorAction Stop
    Install-Module -Name MSAL.PS -Force
    Import-Module -Name MSAL.PS -ErrorAction Stop
}


################################### Users ###################################

. .\Credentials.ps1

$Headers = @{
    'Authorization' = "Bearer $($AuthResponse.AccessToken)"
}


do {
    $ProjectId = Read-Host -Prompt 'Enter Project ID [000]'
    $ProjectIdInt = [int] ($ProjectId)
} while (($ProjectId -eq "") -or !($ProjectIdInt -is [int]))

$ProjectId = [Math]::Truncate($ProjectId)
$ProjectId = '{0:d3}' -f [int]$ProjectId

do {
    $GroupAdminEmail = Read-Host -Prompt 'Enter Project Admin-Email'
    try {
        $GroupAdminUser = Get-AzureADUser -ObjectId $GroupAdminEmail
    }
    catch [Microsoft.Open.AzureAD16.Client.ApiException] {
        $GroupAdminUser = "404"
        Write-Host "User not exists" -ForegroundColor Red
    }
} while (($GroupAdminEmail -eq "") -or ($GroupAdminUser -eq "404"))


Write-Host "Creating Groups for 'Project-$ProjectId'" -ForegroundColor Green

$Confirm = Read-Host -Prompt 'Confirm creation of groups [y] / [n]'

if ($Confirm -eq 'y') {

    $ProjectName = $("Project-" + $ProjectId)

    $ProjectBoardsGroup = Get-AzureADGroup -SearchString "ProjectBoards_Users"

    $All = New-AzureADGroup -DisplayName $("Project-" + $ProjectId + "__All") -Description $("Project-" + $ProjectId + "__All") -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
    $Workers = New-AzureADGroup -DisplayName $("Project-" + $ProjectId + "__Workers") -Description $("Project-" + $ProjectId + "__Workers") -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
    $Externals = New-AzureADGroup -DisplayName $("Project-" + $ProjectId + "__Externals") -Description $("Project-" + $ProjectId + "__Externals") -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"

    # Add groups to parent group
    Add-AzureADGroupMember -ObjectId $All.ObjectId -RefObjectId $Workers.ObjectId
    Add-AzureADGroupMember -ObjectId $All.ObjectId -RefObjectId $Externals.ObjectId

    # add group to project boards users
    Add-AzureADGroupMember -ObjectId $ProjectBoardsGroup.ObjectId -RefObjectId $All.ObjectId

    # Create M365 Group
    $M365Group = New-AzureADMSGroup -DisplayName $("Project-" + $ProjectId) -Description $("Project-" + $ProjectId) -MailEnabled $True -MailNickname $("project-" + $ProjectId) -SecurityEnabled $True -GroupTypes "Unified"

    Write-Host "Waiting for groups to apply" -NoNewline

    $timeout = [Diagnostics.Stopwatch]::StartNew()
    while (($timeout.Elapsed.TotalSeconds -lt 120)) {
        Start-Sleep 0.6
        Write-Host "." -NoNewline
    }
    $timeout.Stop()

    # add owners to the groups
    Add-AzureADGroupOwner -ObjectId $All.ObjectId -RefObjectId $GroupAdminUser.ObjectId
    Add-AzureADGroupMember -ObjectId $All.ObjectId -RefObjectId $GroupAdminUser.ObjectId
    Add-AzureADGroupOwner -ObjectId $Workers.ObjectId -RefObjectId $GroupAdminUser.ObjectId
    Add-AzureADGroupMember -ObjectId $Workers.ObjectId -RefObjectId $GroupAdminUser.ObjectId
    Add-AzureADGroupOwner -ObjectId $Externals.ObjectId -RefObjectId $GroupAdminUser.ObjectId
    Add-AzureADGroupMember -ObjectId $Externals.ObjectId -RefObjectId $GroupAdminUser.ObjectId
    Add-AzureADGroupOwner -ObjectId $M365Group.Id -RefObjectId $GroupAdminUser.ObjectId
    Add-AzureADGroupMember -ObjectId $M365Group.Id -RefObjectId $GroupAdminUser.ObjectId


    Write-Host "Group Administrator $($GroupAdminUser.DisplayName) <$($GroupAdminUser.UserPrincipalName)> can now add project members here 'https://account.activedirectory.windowsazure.com/redirect/groups'" -ForegroundColor Yellow

    Write-Host
    Write-Host
    Write-Host

    ################################### Planner ###################################

    Write-Host "Creating MS Planner Schema"


    #### Planner Boards Creation ####



    $Body = @{
        "owner" = $M365Group.Id
        "title" = $("Project-" + $ProjectId)
    }

    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/planner/plans'
        'Method'      = 'POST'
        'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }

    $Result = Invoke-RestMethod @Params -Headers $Headers

    $PlannerPlanId = $Result.id

    Write-Host "PLAN CREATED"
    Write-Host $Result
    Write-Host 





    $PlannerBuckets = @( "1-Project Preparation", "2-Initialization", "3-Concept", "4-Realization", "5-Preparation Going Live", "6-Going Live" )

    foreach ($Bucket in $PlannerBuckets) {


        $Body = @{
            "name"      = $Bucket
            "planId"    = $PlannerPlanId
            "orderHint" = " !"
        }

        $Params = @{
            'Uri'         = 'https://graph.microsoft.com/v1.0/planner/buckets'
            'Method'      = 'POST'
            'Body'        = $($Body | ConvertTo-Json -Depth 5)
            'ContentType' = 'application/json'
        }

        $Result = Invoke-RestMethod @Params -Headers $Headers

        Write-Host "BUCKET: $Bucket"
        Write-Host $Result
        Write-Host 



    }

    Write-Host "Buckets created"
    Write-Host


    ################################### MS Teams ###################################

    Write-Host "Creating MS Teams"
    Write-Host

    $Body = @{
        "memberSettings"    = @{
            "allowCreateUpdateChannels"         = $true
            "allowDeleteChannels"               = $true
            "allowAddRemoveApps"                = $true
            "allowCreateUpdateRemoveTabs"       = $true
            "allowCreateUpdateRemoveConnectors" = $true
        }
        "guestSettings"     = @{
            "allowCreateUpdateChannels" = $false
            "allowDeleteChannels"       = $false
        }
        "messagingSettings" = @{
            "allowUserEditMessages"    = $false
            "allowUserDeleteMessages"  = $false
            "allowOwnerDeleteMessages" = $true
            "allowTeamMentions"        = $true
            "allowChannelMentions"     = $true
        }
        "funSettings"       = @{
            "allowGiphy"            = $false
            "giphyContentRating"    = "strict"
            "allowStickersAndMemes" = $false
            "allowCustomMemes"      = $false
        }
    }

    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/groups/' + $M365Group.Id + '/team'
        'Method'      = 'PUT'
        'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }

    $Result = Invoke-RestMethod @Params -Headers $Headers


    Write-Host "Successful created MS Teams"



    $Body = @{}

    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/groups'
        'Method'      = 'GET'
        # 'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }
    
    $Result = Invoke-RestMethod @Params -Headers $Headers
    
    $TeamsId = ""
    foreach ($group in $Result.value) {
        if ($group.displayname -eq $ProjectName ) {
            $TeamsId = $group.id
        }
    }
    Write-Host "TeamsId: $TeamsId" -ForegroundColor Cyan
    
    
    
    $Body = @{}
    
    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/teams/' + $TeamsId + '/channels'
        'Method'      = 'GET'
        # 'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }
    
    $Result = Invoke-RestMethod @Params -Headers $Headers
    
    $ChannelId = ""
    foreach ($channel in $Result.value) {
        if ($channel.displayName -eq "General") {
            $ChannelId = $channel.id
        }
    }
    Write-Host "ChannelId: $ChannelId" -ForegroundColor Cyan
    Write-Host
    
    
    # Add Planner Tab to Teams
    
    $EmbeddedPlannerURL = 'https://tasks.office.com/microsoft365devch.onmicrosoft.com/Home/PlannerFrame?page=7&planId=' + $PlannerPlanId
    
    $Body = @{
        "name"                = $ProjectName + " Tasks"
        "displayName"         = $ProjectName + " Tasks"
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner"
        "configuration"       = @{
            "entityId"   = $EmbeddedPlannerURL
            "contentUrl" = $EmbeddedPlannerURL
            "removeUrl"  = $EmbeddedPlannerURL
            "websiteUrl" = $EmbeddedPlannerURL
        }
    }
    
    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/teams/' + $TeamsId + '/channels/' + $ChannelId + '/tabs'
        'Method'      = 'POST'
        'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }
    
    $Result = Invoke-RestMethod @Params -Headers $Headers
    
    # Add OneNote Tab

    $Body = @{}
    
    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/groups/' + $M365Group.Id + '/onenote/notebooks'
        'Method'      = 'GET'
        # 'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }
    
    $Result = Invoke-RestMethod @Params -Headers $Headers
    
    $OneNoteId = $Result.value.id
    $OneNoteLink = $Result.value.links.oneNoteWebUrl.href

    $Body = @{
        "name"                = $ProjectName + " Notes"
        "displayName"         = $ProjectName + " Notes"
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/0d820ecd-def2-4297-adad-78056cde7c78"
        "configuration"       = @{
            "entityId"   = $OneNoteLink
            "contentUrl" = $OneNoteLink
            "removeUrl"  = $OneNoteLink
            "websiteUrl" = $OneNoteLink
        }
    }
    
    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/teams/' + $TeamsId + '/channels/' + $ChannelId + '/tabs'
        'Method'      = 'POST'
        'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }
    
    $Result = Invoke-RestMethod @Params -Headers $Headers
    

    # SharePoint Libraries

    $Body = @{}
    
    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/groups/' + $M365Group.Id + '/sites/root/'
        'Method'      = 'GET'
        # 'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }
    
    $Result = Invoke-RestMethod @Params -Headers $Headers
    
    $SharePointId = $Result.id
    $SharePointLink = $Result.webUrl
    
    
    # Document Library
    
    $Body = @{
        "displayName" = $ProjectName + "-Files"
        "list"        = @{
            "contentTypesEnabled" = $false
            "hidden"              = $false
            "template"            = "documentLibrary"
        }
    }
    
    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/sites/' + $SharePointId + '/lists'
        'Method'      = 'POST'
        'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }
    
    $Result = Invoke-RestMethod @Params -Headers $Headers
    
    $SPLibraryId = $Result.id
    $SPLibraryLink = $Result.webUrl



    $Body = @{
        "name"                = $ProjectName + " SharePoint"
        "displayName"         = $ProjectName + " SharePoint"
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.web"
        "configuration"       = @{
            "entityId"   = $SharePointLink
            "contentUrl" = $SharePointLink
            "removeUrl"  = $SharePointLink
            "websiteUrl" = $SharePointLink
        }
    }

    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/teams/' + $TeamsId + '/channels/' + $ChannelId + '/tabs'
        'Method'      = 'POST'
        'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }

    $Result = Invoke-RestMethod @Params -Headers $Headers



    # SP Library Teams
    $Body = @{
        "name"                = $ProjectName + " DeliveryFiles"
        "displayName"         = $ProjectName + " DeliveryFiles"
        "teamsApp@odata.bind" = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/com.microsoft.teamspace.tab.files.sharepoint"
        "configuration"       = @{
            "entityId"   = ""
            "contentUrl" = $SPLibraryLink
            "removeUrl"  = $null
            "websiteUrl" = $null
        }
    }

    $Params = @{
        'Uri'         = 'https://graph.microsoft.com/v1.0/teams/' + $TeamsId + '/channels/' + $ChannelId + '/tabs'
        'Method'      = 'POST'
        'Body'        = $($Body | ConvertTo-Json -Depth 5)
        'ContentType' = 'application/json'
    }

    $Result = Invoke-RestMethod @Params -Headers $Headers






    Write-Host 'Program end'
    Write-Host

}
else {
    Write-Host 'exit' -ForegroundColor Cyan
    exit
}
