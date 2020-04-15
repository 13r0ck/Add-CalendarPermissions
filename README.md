# Add-Calendar Permissions
## Simplify modifying calender permisisons in Office 365

# Why use this module
* Modify user permission for calendars without having to follow a guide every time
* One-to-many or Many-to-one calendar permssion modifications with ease!
* Specify calendar permissions

For some reason the Office 365 admin protal does not have the option to modify calendar permissions. Maning that to add a user to a calendar each user must do it manually if you do not use powershell. I like PowerShell, others don't. This is for those who do not. Just run the command, login, provide usernames and then the script will take care of the rest of the settings changes for you.

# How to
`Add-Calendar Permissions` for modifying calendar permissions for one user
`Add-Calendar Permissions -AllUser` for modifying one-to-many or many-to-one calendar permisison changes. Such as giveing the CEO acccess to everyone's calendar, or adding everyone to the conference room calendar, respectively.
Then multiple dialogue messages will hold your hand to give the required information to finish the permission changes.

# Install Instructions
This module is not yet signed and uploaded to PowerShell Gallery, for now just copy the folder from this repository to a PowerShell Module folder.
