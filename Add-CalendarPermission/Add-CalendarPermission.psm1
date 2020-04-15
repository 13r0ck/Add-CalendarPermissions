function Add-CalendarPermission {
Param([switch]$AllUsers)
    #Setup Variables
    $TextInfo = (Get-Culture).TextInfo



    if ($null -eq (Get-PSSession | where {$_.ComputerName -eq "outlook.office365.com"})) {
        try {
            Connect-ExchangeOnline | Import-Module -Global
        } catch {
            Throw "It does not appear that you have the correct dependencies for this modules.`nFIX:Run ''Install-Module -Name ExchangeOnlineManagement'' in powershell to install the required dependencies`nOr run Update-AmnetModules`n`nto verify that you have all dependencies for the newest release."
        }
    }

   #The ExhangeOnlineManagement is still under development. This is so that hopefully in the future it will get an update, and someone will notice when running this script to let me know to update the script to use the new faster modules
   if ($null -ne (Get-Command -Name "Set-EXOMailboxFolderPermission" -ea 0)) {Write-Output "`n`n It appears that the ExchangeOnlineManagement module has been updated, but this module has not been. Please contact the maintainer of this script to let them know to update this to the newer faster moduels. Thanks! `n`n"}

   $close_maybe = $null
    while ("$close_maybe".ToUpper() -ne "N") { #Let the technician enter usernames one by one.


        if ($AllUsers) { #If a list of users was piped into the
            Write-Verbose "Grabbing all Users"
            $get_users = Get-User
            foreach ($user in $get_users) {
                if (($user.RecipientType -eq "UserMailbox") -and !($user.accountdisabled)) { #verifiy that a mailbox is used
                    [System.Array]$users += $user.Name #Just create an easy to use array if strings
                }
            }
            #$input
            Write-Host "A list of users were input, is that list:"
            switch (Read-Host "     (1) Owner(s) of calendar(s) giving permission to a specific user.`n     (2) User(s) getting access to a specific calendar`n") { #Makes the code more readable + switches the order in final loop.
                #############################################################################
                ##### - Owner(s) of calendar(s) giving permission to a specific user. - #####
                #############################################################################
                "1" { 
                        $owners_of_calendars = $users
                        $give_permissions_to = Read-Host "Username of who will be getting permissions to multiple calendars"
                        if ($pass -ne 1)
                        {
@"
    `n`nPossible Access rights are:
    Owner — read, create, modify and delete all items and folders. Also this role allows manage items permissions;
    PublishingEditor — read, create, modify and delete items/subfolders;
    Editor — read, create, modify and delete items;
    PublishingAuthor — read, create all items/subfolders. You can modify and delete only items you create;
    Author — create and read items; edit and delete own items NonEditingAuthor – full read access and create items. You can delete only your own items;
    Reviewer — read-only;
    Contributor — create items and folders;
    AvailabilityOnly — read free/busy information from the calendar;
    LimitedDetails;
    None — no permissions to access folder and files.`n
"@
                        }
                        $access_rights = Read-Host "What permissions will be added to the calendars?"
                        @"
`n`nThe following type of action will be performed. (Multiple times if a list of users was passed to this Module)

        $("$($give_permissions_to)".ToUpper()) will gain $($TextInfo.ToTitleCase("$access_rights")) permission to $("$($owners_of_calendars[0])".ToUpper())'s calendar

"@
                        start-sleep -s 0.5
                        if ((Read-Host "Would you like to perform this action? (Y) Yes (N) No. Default is Yes").ToUpper() -eq "N")
                        { #confirm action.
                            break
                        }
                        else
                        {
                            for ($i=0;$i -lt $users.Length;$i++) #($user in $owners_of_calendars)
                            {
                            $user = $users[$i]
                                try
                                {
                                    try
                                    {
                                        Add-MailboxFolderPermission -Identity "$($user):\calendar" -user $give_permissions_to -AccessRights $access_rights -ErrorAction Stop | Out-Null
                                        "Adding $give_permissions_to $access_rights permissions to $user's calendar"
                                    }
                                    catch
                                    {
                                        Set-MailboxFolderPermission -Identity "$($user):\calendar" -user $give_permissions_to -AccessRights $access_rights -ErrorAction Stop | Out-Null
                                        "Changing $give_permissions_to permissions to $access_rights for $user's calendar"
                                    }
                                }
                                catch
                                {
                                    Write-Error "Either $user or $give_permissions_to is not a valid username, or $user does not have a calendar."
                                    $there_were_errors = $true
                                }
                            } #End for each
                        }
                    }
                #############################################################
                ##### - User(s) getting access to a specific calendar - #####
                #############################################################
                "2" {
                        $global:users_gaining_access = $users
                        $global:shared_calendar = Read-Host "Username associated with the calendar that is to be shared"
                        if ($pass -ne 1)
                        {
@"
    `n`nPossible Access rights are:
    Owner — read, create, modify and delete all items and folders. Also this role allows manage items permissions;
    PublishingEditor — read, create, modify and delete items/subfolders;
    Editor — read, create, modify and delete items;
    PublishingAuthor — read, create all items/subfolders. You can modify and delete only items you create;
    Author — create and read items; edit and delete own items NonEditingAuthor – full read access and create items. You can delete only your own items;
    Reviewer — read-only;
    Contributor — create items and folders;
    AvailabilityOnly — read free/busy information from the calendar;
    LimitedDetails;
    None — no permissions to access folder and files.`n
"@
                        }
                        $access_rights = Read-Host "What permissions will be given to the users for the calendar?"
@"
`n`nThe following type of action will be performed. (Multiple times if a list of users was passed to this Module)

        $("$($users_gaining_access[1])".ToUpper()) will gain $($TextInfo.ToTitleCase("$access_rights")) permission to $("$($shared_calendar)".ToUpper())'s calendar

"@
                        start-sleep -s 0.5
                        if ((Read-Host "Would you like to perform this action? (Y) Yes (N) No. Default is Yes").ToUpper() -eq "N") { #confirm action.
                            break
                        }
                        else
                        {
                            for($i=0;$i -lt $users.Length; $i++)# ($give_permisions_to in $global:users_gaining_access)
                            {
                                $give_permissions_to = $users[$i]
                                try
                                {
                                    try
                                    {
                                        Add-MailboxFolderPermission -Identity "$($global:shared_calendar):\calendar" -user $give_permissions_to -AccessRights $access_rights -ErrorAction Stop
                                        "Adding $give_permissions_to $access_rights permissions to $global:shared_calendar's calendar"
                                    }
                                    catch
                                    {
                                        Set-MailboxFolderPermission -Identity "$($global:shared_calendar):\calendar" -user $give_permissions_to -AccessRights $access_rights -ErrorAction Stop
                                        "Changing $give_permissions_to permissions to $access_rights for $global:shared_calendar's calendar"
                                    }
                                }
                                catch
                                {
                                    "Error $_"
                                    Write-Error "Either $global:shared_calendar or $give_permissions_to is not a valid username, or $global:shared_calendar does not have a calendar."
                                    $there_were_errors = $true
                                }
                            } #end foreach
                        }
                    }
                default {Throw "$("'" + $_ + "'" + " is not a valid input.")"} 
            }

        ################################
        ##### - Enter one by one - #####
        ################################
        }
        else
        { #The -allusers switch was not used
            $owners_of_calendars = @()
            $user = Read-Host "Username of the user who owns the calendar"
            $give_permissions_to = Read-Host "Username of who will be getting access to $user's calendar"
            $example_user = $user
        
            if ($pass -ne 1)
            {
@"
    `n`nPossible Access rights are:
    Owner — read, create, modify and delete all items and folders. Also this role allows manage items permissions;
    PublishingEditor — read, create, modify and delete items/subfolders;
    Editor — read, create, modify and delete items;
    PublishingAuthor — read, create all items/subfolders. You can modify and delete only items you create;
    Author — create and read items; edit and delete own items NonEditingAuthor – full read access and create items. You can delete only your own items;
    Reviewer — read-only;
    Contributor — create items and folders;
    AvailabilityOnly — read free/busy information from the calendar;
    LimitedDetails;
    None — no permissions to access folder and files.`n
"@
            }
            $access_rights = Read-Host "What permissions will be given to $give_permissions_to"
@"
`n`nThe following type of action will be performed. (Multiple times if a list of users was passed to this Module)

        $("$give_permissions_to".ToUpper()) will gain $($TextInfo.ToTitleCase("$access_rights")) permission to $("$user".ToUpper())'s calendar

"@
        start-sleep -s 0.5
        if ((Read-Host "Would you like to perform this action? (Y) Yes (N) No. Default is Yes").ToUpper() -eq "N")
        { #confirm action.
            break
        }
        else
        {
            try
            {
                try
                {
                    Add-MailboxFolderPermission -Identity "$($user):\calendar" -user $give_permissions_to -AccessRights $access_rights -ErrorAction Stop | Out-Null
                    "Adding $give_permissions_to $access_rights permissions to $user's calendar"
                }
                catch
                {
                    Set-MailboxFolderPermission -Identity "$($user):\calendar" -user $give_permissions_to -AccessRights $access_rights -ErrorAction Stop | Out-Null
                    "Changing $give_permissions_to permissions to $access_rights for $user's calendar"
                }
            }
            catch
            {
                Write-Error "Either $user or $give_permissions_to is not a valid username, or $user does not have a calendar."
                $there_were_errors = $true
            }
        }
        $close_maybe = read-host "Another User? (Y) Yes (N) No"
        $pass = 1
        if ($AllUsers) {break}
    } #End the Y/N for more input loop
    if ($there_were_errors)
    {
        Write-Host "If you did not notice, some accounts did have errors while modifying permissions. Please scroll up to read more."
    }
}
}
Export-ModuleMember -Function Add-CalendarPermission