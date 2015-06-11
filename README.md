# Office365
Office 365 Auto-Provision Script

This is an Office 365 PowerShell Auto-Provisioning script that can be run as a scheudled task in conjuction with your AD syncrhonization process.

# Usage

You will need to modify certain aspects of the script to fit your environment.  Curerently it is setup to used a stored credential to connect to the Azure Cloud/MSONline service.

The script will pull uses from a specified group in AD, match that list against against MSOnline and verify license count and the "isLicensed" attribute to see if user is licenses or not.  If licenses are availabe a license is issue, if not the lciense count logged to the logFile.
