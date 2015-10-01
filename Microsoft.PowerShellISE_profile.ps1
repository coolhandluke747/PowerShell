<#
    Last Updated
    October 1st, 2015
    Version 2.2
    
    .SYNOPSIS
    PowerShell Script to Connect to Office 365 Services & On-Prem Services

    PowerShell script defining functions to connect to Office 365 online services 
    and Exchange & Lync On-Premises. 

    Revision History
    ---------------------------------------------------------------------
    2.0 - Updated the add-ons menu and added more admin functions.
    2.1 - Moved the functions around to make it more readable.
    2.2 - Added setCorrectTabNames to see what your connected to if you connect to different services at different times.

    .INSTALLATION INSTRUCTIONS
    1. Update PowerShell to Latest Version - (Major 4) as of this writing.
    You can check version by running this command "$PSVersionTable.PSVersion"

    2. Install all necessary PowerShell Modules 32-bit or 64-bit depending
    on the computer to connect to the various services of Office 365.
    
    3. Run: "Set-ExecutionPolicy RemoteSigned"

    4. Place this file in C:\Users\<USERNAME>\Documents\WindowsPowershell

    5. Restart PowerShellISE or open a new tab to see the changes take effect

    .DESCRIPTION 
    The functions are listed below. 
    
    - ConnectAll                         Connects to all Office 365 Services and on-Prem Exchange and Lync          		
    - ConnectO365                        Connects to Office 365         		
    - ConnectEXO                         Connects to Exchange Online
    - ConnectSPO                         Connects to SharePoint Online        		
    - ConnectSFBO                        Connects to Skype for Business Online         		
    - ConnectCC   		                 Connects to Compliance Center
    - ConnectAzureRMS                    Connects to Azure Rights Management
    - ConnectExchange                    Connects to on-prem Exchange Server
    - ConnectLync	                     Connect to on-prem Lync Server
    - OpenIEAdmin                        Gives permission to mailbox as username@address.com and opens mailbox in IE
    - OpenChromeAdmin                    Gives permission to mailbox as username@address.com and opens mailbox in Chrome
    - OpenIEEmail                        Gives permission to mailbox as username@address.com or other specified mailbox and opens mailbox in IE 
    - OpenChromeEmail                    Gives permission to mailbox as username@address.com or other specified mailbox and opens mailbox in Chrome
    - Get-UserInfo                       Gets users information on browser, devices, etc by default for 3 days or other if specified
    - GrantPermissionToAllMailboxes      Grants username@address.com or other if specified access to all mailboxes in tenant
    - RemovePermissionToAllMailboxes     Removes username@address.com or other if specified access to all mailboxes in tenant
    - RemovePermissionToMailbox          Removes username@address.com or other if specified access to a specific mailbox
    - CheckForwarding                    Checks forwarding and gives the option to change it to something else or nothing
    - prompt                             Change how the prompt looks  
    - Format-Ticket                      Formats TD Tickets
    - Get-Any                            Try and find if there is anything with the email address
    - Convert-X500                       Converts a stale contact into the corrisponding X500 address
    - o365BoxExists                      Checks if a box exists
    - isConnectedTo                      Tests if a remote session is open with a given computer
    - disconnectAll                      Disconnects all remote sessions and exits all sessions
    - get-Cred                           Uses a encrypted file to store credential information
    - set-TabName                        Changes the display and the prompt
    - setCorrectTabNames                 Checks each service and if multiple services are connected it combines the tab name to what's connected
    - initMenu                           Creates the menu in add-ons
    
#>

cd "$env:HOMEDRIVE$env:HOMEPATH\Documents\WindowsPowerShell"

$FormatEnumerationLimit = -1

$Global:prompt = "Local"
#############################################################
# Custom Functions
# These get defined every time powershell ISE opens
#############################################################

# Change how the prompt looks
function prompt {"PS $Global:prompt> "}

# Formats TD Tickets
Function Format-Ticket
{
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$Unformatted)

    $Formatted = $Unformatted -replace "      ", "`n"
    
    while($Formatted -match "`n ")
    {
        $Formatted = $Formatted -replace "`n ", "`n"
    }

    $Formatted
}

# Try and find if there is anything with the email address
function Get-Any($email)
{
    $countErr = 0
    Get-Mailbox $email 2>&1 | Out-Null
    if ($? -eq $false){$countErr++}else{Write-host "Mailbox"; Get-Mailbox $email}

    Get-MailUser $email 2>&1 | Out-Null
    if ($? -eq $false){$countErr++}else{Write-host "Mail User"; Get-MailUser $email}

    Get-MailContact $email 2>&1 | Out-Null
    if ($? -eq $false){$countErr++}else{Write-host "Mail Contact"; Get-MailContact $email}

    if ($countErr -eq 3){"Does not exist"}
}

# Converts a stale contact into the corrisponding X500 address
Function Convert-X500
{
    $ADDR= READ-HOST "ENTER FULL IMCEAEX ADDRESS:" 
    $REPL= @(@("_","/"), @("\+20"," "), @("\+28","("), @("\+29",")"), @("\+2C",","), @("\+5F", "_" ), @("\+40", "@" ), @("\+2E", "." )) 
    $REPL | FOREACH { $ADDR= $ADDR -REPLACE $_[0], $_[1] } 
    $ADDR= "X500:$ADDR" -REPLACE "IMCEAEX-","" -REPLACE "@.*$", "" 
    WRITE-HOST $ADDR
}

#############################################################
# Helper Functions
# Common tasks that are done often, used to add readablity 
#############################################################
# Checks if a box exists
function o365BoxExists($box)
{
    if(@(Get-Mailbox $box).count -eq 1){return $true}else{return $false}
}

# Tests if a remote session is open with a given computer
function isConnectedTo($connection)
{
    $session = Get-PSSession 
    if(($session.ComputerName -like $connection) -and ($session.state -eq "Opened")){return $true}else{return $false}
}

# Disconnects all remote sessions and exits all sessions
function disconnectAll($test)
{
    Exit-PSSession
	Get-PSSession | Remove-PSSession
    set-TabName(setCorrectTabNames)
}

# Uses a encrypted file to store credential information
function get-Cred
{
    Param([string]$credPath,
          [string]$account)

    if ((Test-Path $credPath) -eq $true)
    {
        $password = cat $credPath | convertto-securestring
        $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $account, $password
    }
    else
    {
        $creds = Get-Credential –credential $account
        $encp = $creds.password 
        $encp | ConvertFrom-SecureString | Set-Content $credPath

        $password = cat $credPath | convertto-securestring
        $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $account, $password
    }

    return $cred
}

# Changes the display and the prompt
function set-TabName($newName)
{
    
    $i = 0
    $psISE.PowerShellTabs | foreach{
        if (($_.DisplayName -match "$newName*") -or ($_.DisplayName -eq $newName))
        {
            $i++
        }
    }

    if ($i -eq 0)
    {
        $i = ""
    }
    else
    {
        $i = " ($i)"
    }

    $Global:prompt = "$newName$i"
    $psISE.CurrentPowerShellTab.DisplayName = "$newName$i"
}

# Checks each service and if multiple services are connected it combines the tab name to what's connected
function setCorrectTabNames
{

    if((isConnectedTo("Your Exchange on-prem address.")) -eq $true)

    {
        $arrayOfConnections = ("Exchange ")
    }

    else

    {
    #Not Connected
    }

    if((isConnectedTo("Your Lync on-prem address")) -eq $true)

    {
        $arrayOfConnections += "Lync "
    }

    else

    {
    #Not Connected
    }

    if((isConnectedTo("outlook.office365.com")) -eq $true)

    {
        $arrayOfConnections += "EXO/EOP "
    }

    else

    {
    #Not Connected
    }


    if(Get-MsolDomain -ErrorAction SilentlyContinue)

    {
        $arrayOfConnections += "O365 "
    }

    else

    {
    #Not Connected
    }

    try
        {
            if(Get-SPOSite)

            {
                $arrayOfConnections += "SPO "
            }

            else

            {
                Get-SPOSite -ErrorAction SilentlyContinue
            }
            
        }
        catch
        {
            #Not Connected   
        }

    
    if((isConnectedTo("admin0b.online.lync.com")) -eq $true)

    {
        $arrayOfConnections += "SFBO "
    }

    else

    {
    #Not Connected
    }


    if((isConnectedTo("ps.compliance.protection.outlook.com")) -eq $true)

    {
        $arrayOfConnections += "CC "
    }

    else

    {
    #Not Connected
    }


    try
        {
            if(Get-Aadrm)

            {
                $arrayOfConnections += "ARM "
            }

            else

            {
                Get-Aadrm -ErrorAction SilentlyContinue
            }
            
        }
        catch
        {
            #Not Connected   
        }


        if($arrayOfConnections.Count -eq 1)
        {
            Return $arrayOfConnections
        }

        else

        {
            Return $arrayOfConnections 
        }
}

#############################################################
# Main Connection Functions
# For connecting to various Office 365 Services.
#############################################################

# Connects to all Office 365 Services and on-Prem Exchange and Lync
function ConnectAll
{

    ConnectExchange
    ConnectLync
    ConnectO365
    ConnectEXO/EOP
    ConnectSPO
    ConnectSFBO
    ConnectCC
    ConnectAzureRMS
     
    Set-TabName(setCorrectTabNames)
}
  
# Connect to on-prem Exchange
function ConnectExchange
{
    
    $username = $env:username
    

    if((isConnectedTo("Your Exchange on-prem address.")) -eq $false)

    {
    Write-Host "Not Connected to Your Exchange on-prem address, Connecting Now as $username" -ForegroundColor Red
    }

    else

    {
    Write-Host "Already Connected to Your Exchange on-prem address as $username" -ForegroundColor Green
    Return
    }

    
    $s = New-PSSession -ConfigurationName Microsoft.Exchange `
    -ConnectionUri "Your Exchange PowerShell on-prem address." `
    -Authentication Kerberos
    Import-PSSession $s -DisableNameChecking -Prefix exh
    
    set-TabName(setCorrectTabNames)
}

# Connect to on-prem Lync Server
function ConnectLync
{

    $username = $env:username
    $path = "C:\Users\$username\Documents\WindowsPowerShell\LyncCredentials.txt"
    $cred = get-Cred -account "username@address.com" -credPath $path
    $connectName = $cred.UserName
    
    if((isConnectedTo("Your Lync on-prem address")) -eq $false)

    {
    Write-Host "Not Connected to Your Lync on-prem, Connecting Now as $connectName" -ForegroundColor Red
    }

    else

    {
    Write-Host "Already Connected to Your Lync on-prem as $connectName" -ForegroundColor Green
    Return
    }
    
    $s = New-PSSession -ConnectionUri "Your Lync PowerShell on-prem address." `
    -Credential $cred

    Import-PSSession $s -Prefix lync
    
    set-TabName(setCorrectTabNames)
}

# Connect to Exchange Online & other functions related to Exchange Online
function ConnectEXO/EOP
{
        $username = $env:username
        $path = "C:\Users\$username\Documents\WindowsPowerShell\Office365Credentials.txt"
        $cred = get-Cred -account "username@address.com" -credPath $path
        $connectName = $cred.UserName

        if((isConnectedTo("outlook.office365.com")) -eq $false)

        {

        Write-Host "Not Connected to Exchange Online/Exchange Online Protection, Connecting Now as $connectName" -ForegroundColor Red

        
        
        @($Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection) 2>&1 | Out-Null
        Import-PSSession $Session -DisableNameChecking
        set-TabName(setCorrectTabNames)


        }

        else

        {
        Write-Host "Already Connected to Exchange Online/Exchange Online Protection as $connectName" -ForegroundColor Green
        Return
        }

          
 }

# Connect to Office 365
function ConnectO365
{
          $username = $env:username
          $path = "C:\Users\$username\Documents\WindowsPowerShell\Office365Credentials.txt"
          $cred = get-Cred -account "username@address.com" -credPath $path
          $connectName = $cred.UserName

          if(Get-MsolDomain -ErrorAction SilentlyContinue)

          {
             Write-Host "Already Connected to Office 365 as $connectName" -ForegroundColor Green
             Return
          }

          else

          {
          Write-Host "Not Connected to Office 365, Connecting Now as $connectName" -ForegroundColor Red

         
          connect-msolservice -Credential $cred
          set-TabName(setCorrectTabNames)

          }
 
}
  
# Connect to SharePoint Online
function ConnectSPO
{

        $username = $env:username
        $path = "C:\Users\$username\Documents\WindowsPowerShell\Office365Credentials.txt"
        $cred = get-Cred -account "username@address.com" -credPath $path
        $connectName = $cred.UserName


        Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking

        try
        {
            if(Get-SPOSite)

            {
                Write-Host "Already Connected to SharePoint Online as $connectName" -ForegroundColor Green
                Return
            }

            else

            {
                Get-SPOSite -ErrorAction SilentlyContinue
            }
            
        }
        catch
        {
            Write-Host "Not Connected to SharePoint Online, Connecting Now as $connectName" -ForegroundColor Red

            #Connect to SPO, you have to connect before anything can be executed.
            Connect-SPOService -Url "Your SharePoint Online Admin URL" -credential $cred

            #Change the tab Name.
            set-TabName(setCorrectTabNames)   
        }
        <#
        Need to figure out a way to use these in a function or something, but above will work just for simple overall SPO server calls.

        Client Side Manipulation Libraries.
        You have to install Client Side Object Model for Powershell to do anything specific for a client in SharePoint like a list, content types, etc.
        References to SharePoint Client Assemblies.
        $SPPath = "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16"
        Add-Type -Path "$SPPath\ISAPI\Microsoft.SharePoint.Client.dll"
        Add-Type -Path "$SPPath\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
        References to Service Application Specific Assemblies.
        Add-Type -Path "$SPPath\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"
        Add-Type -Path "$SPPath\ISAPI\Microsoft.SharePoint.Client.Search.dll"
        Add-Type -Path "$SPPath\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll"

        Configure Site URL and Admin Credentials for Web Authentication against SharePoint sites.
        $SiteURL = "https://domain.sharepoint.com"
        $SPUser = "username@address.com"
        $Password = Read-Host -Prompt "Admin Password" -AsSecureString

        $SPPassword = ConvertTo-SecureString $path -AsPlainText -Force 
        $SPPassword = Get-Content blueCred.txt | ConvertTo-SecureString -Key (1..16)

        Bind to Site Collection
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $SPCreds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SPUser,$SPPassword)
        $Context.Credentials = $SPCreds

        Identify Users in the Site Collection
        $Users = $Context.Web.SiteUsers
        $Context.Load($Users)
        $Context.ExecuteQuery()
        #>
}

# Connect to Skype for Business Online
function ConnectSFBO
{

    $username = $env:username
    $path = "C:\Users\$username\Documents\WindowsPowerShell\Office365Credentials.txt"
    $cred = get-Cred -account "username@address.com" -credPath $path
    $connectName = $cred.UserName

    if((isConnectedTo("admin0b.online.lync.com")) -eq $false)

    {
    Write-Host "Not Connected to Skype for Business Online, Connecting Now as $connectName" -ForegroundColor Red
    }

    else

    {
    Write-Host "Already Connected to Skype for Business Online as $connectName" -ForegroundColor Green
    Return
    }

    Import-Module LyncOnlineConnector

    $lyncSession = New-CsOnlineSession -Credential $cred –OverrideAdminDomain "Your override domain if needed"

    Import-PSSession $lyncSession


    #Change the tab Name.
    set-TabName(setCorrectTabNames)
}

# Connect to the Compliance Center
function ConnectCC
{

    $username = $env:username
    $path = "C:\Users\$username\Documents\WindowsPowerShell\Office365Credentials.txt"
    $cred = get-Cred -account "username@address.com" -credPath $path
    $connectName = $cred.UserName

    if((isConnectedTo("ps.compliance.protection.outlook.com")) -eq $false)

    {
    Write-Host "Not Connected to Compliance Center, Connecting Now as $connectName" -ForegroundColor Red
    }

    else

    {
    Write-Host "Already Connected to Compliance Center as $connectName" -ForegroundColor Green
    Return
    }

    $ccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection

    Import-PSSession $ccSession -Prefix cc -Verbose

    #Change the tab Name.
    set-TabName(setCorrectTabNames)
}

# Connect to Azure Rights Management Service
function ConnectAzureRMS
{

        $username = $env:username
        $path = "C:\Users\$username\Documents\WindowsPowerShell\Office365Credentials.txt"
        $cred = get-Cred -account "username@address.com" -credPath $path
        $connectName = $cred.UserName

          try
        {
            if(Get-Aadrm)

            {
                Write-Host "Already Connected to Azure Rights Management as $connectName" -ForegroundColor Green
                Return
            }

            else

            {
                Get-Aadrm -ErrorAction SilentlyContinue
            }
            
        }
        catch
        {
            Write-Host "Not Connected to Azure Rights Management, Connecting Now as $connectName" -ForegroundColor Red

            Connect-AadrmService -Credential $cred
 
            #Change the tab Name.
            set-TabName(setCorrectTabNames)   
        }

}

#############################################################
# Troubleshooting Functions
#############################################################
# Gives permission to mailbox as username@address.com and opens mailbox in IE
function OpenIEAdmin
{
    ConnectEXO

    $email = Read-Host 'Email Address'
    if((o365BoxExists($email)) -eq $true){
            Write-host "Accessing: $email as username@address.com" -ForegroundColor Green
            Add-MailboxPermission -Identity $email -user username@address.com -AccessRights fullaccess -InheritanceType all -AutoMapping:$false
            Add-RecipientPermission $email -AccessRights SendAs -Trustee username@address.com -Confirm:$false
            $url = "https://outlook.office365.com/owa/$email/"
            $wshell = New-Object -com WScript.Shell
            $wshell.Run("iexplore.exe $url")
            }
            else
            {
            Write-Host "Mailbox `"$email`" does not exist!" -ForegroundColor red
            }

}

# Gives permission to mailbox as username@address.com and opens mailbox in Chrome
function OpenChromeAdmin
{
    ConnectEXO
    
    $email = Read-Host 'Email Address'
    if((o365BoxExists($email)) -eq $true){
        Write-host "Accessing: $email as username@address.com" -ForegroundColor Green
        Add-MailboxPermission -identity $email -user username@address.com -AccessRights fullaccess -InheritanceType all -AutoMapping:$false
        Add-RecipientPermission $email -AccessRights SendAs -Trustee username@address.com -Confirm:$false 
        $url = "https://outlook.office365.com/owa/$email/"
        $wshell = New-Object -com WScript.Shell
        $wshell.Run("chrome.exe $url")
        }
        else
        {
        Write-Host "Mailbox `"$email`" does not exist!" -ForegroundColor red
        }

    
}

# Gives permission to mailbox as username@address.com or other specified mailbox and opens mailbox in IE 
function OpenIEEmail
{
    ConnectEXO

    $email  = Read-Host 'Email Address'
    
    if((o365BoxExists($email)) -eq $true){
                

                $user = "username@address.com"

                $option = Read-Host "Grant $user to $email <1> or other user <2>? (Default is <1> if left empty)"
        

            switch ($option)
            {
                1 {
                    Write-Host "Granting $user access to $email" -ForegroundColor Green
 
                    Add-MailboxPermission -Identity $email -User $user -AccessRights FullAccess -InheritanceType all -AutoMapping:$false
                    Add-RecipientPermission $email -AccessRights SendAs -Trustee $user -Confirm:$false
                    Write-host "Accessing: $email as $user" -ForegroundColor Green
                    $url = "https://outlook.office365.com/owa/$email/"
                    $wshell = New-Object -com WScript.Shell
                    $wshell.Run("iexplore.exe $url") 
                  }

                2 {
                    $username2  = Read-Host 'Other Username'
            
                        if((o365BoxExists($username2)) -eq $true)
                        {
                            Add-MailboxPermission -Identity $email -User $username2 -AccessRights FullAccess -InheritanceType all -AutoMapping:$false
                            Add-RecipientPermission $email -AccessRights SendAs -Trustee $username2 -Confirm:$false
                            Write-host "Accessing: $email as $username2" -ForegroundColor Green
                            $url = "https://outlook.office365.com/owa/$email/"
                            $wshell = New-Object -com WScript.Shell
                            $wshell.Run("iexplore.exe $url") 
                        }

                        else
                        {
                            Write-Host "$username2 doesn't exist!" -ForegroundColor Red
                            Return
                        } 
                  }

                default {
                          Write-Host "Granting $user access to $email" -ForegroundColor Green
  
                          Add-MailboxPermission -Identity $email -User $user -AccessRights FullAccess -InheritanceType all -AutoMapping:$false
                          Add-RecipientPermission $email -AccessRights SendAs -Trustee $user -Confirm:$false
                          Write-host "Accessing: $email as $user" -ForegroundColor Green
                          $url = "https://outlook.office365.com/owa/$email/"
                          $wshell = New-Object -com WScript.Shell
                          $wshell.Run("iexplore.exe $url") 
                        }
                }
            }

            else

            {
              Write-Host "Mailbox `"$name`" does not exist!" -ForegroundColor red
            } 


}

# Gives permission to mailbox as username@address.com or other specified mailbox and opens mailbox in Chrome
function OpenChromeEmail
{

    ConnectEXO

    $email  = Read-Host 'Email Address'
    
    if((o365BoxExists($email)) -eq $true){

            $user = "username@address.com"
        
            $option = Read-Host "Grant $user to $email <1> or other user <2>? (Default is <1> if left empty)"


   switch ($option)
   {
        1 {
            Write-Host "Granting $user access to $email" -ForegroundColor Green
 
            Add-MailboxPermission -Identity $email -User $user -AccessRights FullAccess -InheritanceType all -AutoMapping:$false
            Add-RecipientPermission $email -AccessRights SendAs -Trustee $user -Confirm:$false
            Write-host "Accessing: $email as $user" -ForegroundColor Green
            $url = "https://outlook.office365.com/owa/$email/"
            $wshell = New-Object -com WScript.Shell
            $wshell.Run("chrome.exe $url") 
        }

        2 {
            $username2  = Read-Host 'Other Username'
            
            Write-Host "Granting $username2 access to $email" -ForegroundColor Green


            if((o365BoxExists($username2)) -eq $true)
            {
                Add-MailboxPermission -Identity $email -User $username2 -AccessRights FullAccess -InheritanceType all -AutoMapping:$false
                Add-RecipientPermission $email -AccessRights SendAs -Trustee $username2 -Confirm:$false
                Write-host "Accessing: $email as $username2" -ForegroundColor Green
                $url = "https://outlook.office365.com/owa/$email/"
                $wshell = New-Object -com WScript.Shell
                $wshell.Run("chrome.exe $url") 
            }

            else
            {
                Write-Host "$username2 doesn't exist!" -ForegroundColor Red
                Return
            }

            
 
            
        }

        default {
            Write-Host "Granting $user access to $email" -ForegroundColor Green
 
            Add-MailboxPermission -Identity $email -User $user -AccessRights FullAccess -InheritanceType all -AutoMapping:$false
            Add-RecipientPermission $email -AccessRights SendAs -Trustee $user -Confirm:$false
            Write-host "Accessing: $email as $user" -ForegroundColor Green
            $url = "https://outlook.office365.com/owa/$email/"
            $wshell = New-Object -com WScript.Shell
            $wshell.Run("chrome.exe $url") 
        }
   }


    }

    else

    {
        Write-Host "Mailbox `"$name`" does not exist!" -ForegroundColor red
    }

}

# Gets users information on browser, devices, etc by default for 3 days or other if specified
function Get-UserInfo
{
    ConnectEXO

    $email  = Read-Host 'Email Address or Username of User'
    $days = Read-Host "How many days do you want a report back from today? (Default is 3 days if left empty)"
    
    if (!$email)
            {
                Write-Host -Message "Email address or username required" -ForegroundColor red
                Return
            }

            else
            {

                $box = get-mailbox $email

            }

            
            if ($days)
            {
                $startdate = $(($(Get-Date).AddDays([int]$days * -1)).ToString("MM/dd/yyyy"))
            }

            else 
            {
                $days = 3
                $startdate = $(($(Get-Date).AddDays($days * -1)).ToString("MM/dd/yyyy"))
            }

            $CurrentDate = Get-Date

            $email = $box.PrimarySmtpAddress
            $username = $box.UserPrincipalName
            $displayName = $box.DisplayName

            $report = "<h1>Report For User: $displayName for past(s): $days days since $CurrentDate</h1>"
            $report += "<b>Username: $username</b><br>"
            $report += "<b>Email: $email</b><br><br>"
            $stats = Get-MailboxStatistics -Identity $username | select ItemCount, LastLogonTime
            $report += $box | select DisplayName, WhenCreated, WhenChanged, @{Expression={"$($stats.ItemCount)"};Label="Item Count"}, @{Expression={"$($stats.LastLogonTime)"};Label="Last Logon"}, @{Expression={if($_.ForwardingSmtpAddress){True}else{False}};Label="Is Forwarding"}, ForwardingSmtpAddress | ConvertTo-Html

            
            $OS = Get-O365ClientOSDetailReport -WindowsLiveID $username -StartDate $startdate

            $Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
"@
            $report += "<h2>Operating Systems Used:</h2>"
            $report += $OS | Select @{Expression={"$($_.Name) $($_.version)"};Label="OS"}, @{Expression={$_.Count};Label="Use Count"}, @{Expression={$_.LastAccessTime};Label="Last Used"} | ConvertTo-HTML -Head $Header

            $Browser = Get-O365ClientBrowserDetailReport -WindowsLiveID $username -StartDate $startdate
            $report += "<b>Browsers Used:</b>"
            $report += $Browser | select @{Expression={$_.Name};Label="Browser"}, @{Expression={$_.Version};Label="Version"}, @{Expression={$_.Count};Label="Use Count"}, @{Expression={$_.LastAccessTime};Label="Last Used"} | ConvertTo-HTML
    

            $report += "<h2>Phones Syncing:</h2>"
            $phones = Get-MobileDeviceStatistics -Mailbox $username
            $report += $phones | select FriendlyName, ClientVersion, DeviceModel, DeviceType, FirstSyncTime, LastSuccessSync, LastSyncAttemptTime | ConvertTo-HTML
    
            $Desktop = [Environment]::GetFolderPath("Desktop")
    
            $report | Out-File $Desktop\UserInfo.html
            Invoke-Item $Desktop\UserInfo.html
}

# Grants username@address.com or other if specified access to all mailboxes in tenant
function GrantPermissionToAllMailboxes
{
    ConnectEXO

    $user  = Read-Host "Who do you want to give all Read & Send Access to all Mailboxes? (Default is username@address.com if left empty)"
    
    if ($user)
            {
                $user = $user
            }

            else
            {
                $user = "username@address.com"
            }

            $startTime = Get-Date
            

            Write-Host "Make sure you run this function in a new PowerShell Tab and/or different computer if you want to do other work, as this will take a long time and lots of computer resources." -ForegroundColor Red

            Write-host "Granting: $user to all Mailboxes started on" $startTime -ForegroundColor Green

            
            Write-host "Gathering all Mailboxes through Display Name of A through F" -ForegroundColor Green

            $mailboxesAF = Get-Mailbox -ResultSize unlimited | ?{$_.DisplayName -match "^[A-F]"}
            
            $counter = 0
            foreach ($mbx in $mailboxesAF){
                $counter++
                $display = $mbx.DisplayName
                Add-MailboxPermission -Identity $mbx.UserPrincipalName -User $user -AccessRights fullaccess -AutoMapping:$false -InheritanceType all
                Add-RecipientPermission -Identity $mbx.UserPrincipalName -AccessRights SendAs -Trustee $user -Confirm:$false
                Write-Progress -Activity "Adding $user read & send access to: $display" -Status "Processing $($counter) of $($mailboxesAF.count)" -CurrentOperation $mbx.WindowsEmailAddress -PercentComplete (($counter / $mailboxesAF.count) * 100)
                Start-Sleep -Milliseconds 200
            }


            Write-host "Gathering all Mailboxes through Display Name of G through L" -ForegroundColor Green

            $mailboxesGL = Get-Mailbox -ResultSize unlimited | ?{$_.DisplayName -match "^[G-L]"}
            
            $counter1 = 0
            foreach ($mbx in $mailboxesGL){
                $counter1++
                $display = $mbx.DisplayName
                Add-MailboxPermission -Identity $mbx.UserPrincipalName -User $user -AccessRights fullaccess -AutoMapping:$false -InheritanceType all
                Add-RecipientPermission -Identity $mbx.UserPrincipalName -AccessRights SendAs -Trustee $user -Confirm:$false
                Write-Progress -Activity "Adding $user read & send access to: $display" -Status "Processing $($counter1) of $($mailboxesGL.count)" -CurrentOperation $mbx.WindowsEmailAddress -PercentComplete (($counter1 / $mailboxesGL.count) * 100)
                Start-Sleep -Milliseconds 200
            }


            Write-host "Gathering all Mailboxes through Display Name of M through R" -ForegroundColor Green

            $mailboxesMR = Get-Mailbox -ResultSize unlimited | ?{$_.DisplayName -match "^[M-R]"}
            
            $counter2 = 0
            foreach ($mbx in $mailboxesMR){
                $counter2++
                $display = $mbx.DisplayName
                Add-MailboxPermission -Identity $mbx.UserPrincipalName -User $user -AccessRights fullaccess -AutoMapping:$false -InheritanceType all
                Add-RecipientPermission -Identity $mbx.UserPrincipalName -AccessRights SendAs -Trustee $user -Confirm:$false
                Write-Progress -Activity "Adding $user read & send access to: $display" -Status "Processing $($counter2) of $($mailboxesMR.count)" -CurrentOperation $mbx.WindowsEmailAddress -PercentComplete (($counter2 / $mailboxesMR.count) * 100)
                Start-Sleep -Milliseconds 200
            }
            

            Write-host "Gathering all Mailboxes through Display Name of S through Z" -ForegroundColor Green

            $mailboxesSZ = Get-Mailbox -ResultSize unlimited | ?{$_.DisplayName -match "^[S-Z]"}
            
            $counter3 = 0
            foreach ($mbx in $mailboxesSZ){
                $counter3++
                $display = $mbx.DisplayName
                Add-MailboxPermission -Identity $mbx.UserPrincipalName -User $user -AccessRights fullaccess -AutoMapping:$false -InheritanceType all
                Add-RecipientPermission -Identity $mbx.UserPrincipalName -AccessRights SendAs -Trustee $user -Confirm:$false
                Write-Progress -Activity "Adding $user read & send access to: $display" -Status "Processing $($counter3) of $($mailboxesSZ.count)" -CurrentOperation $mbx.WindowsEmailAddress -PercentComplete (($counter3 / $mailboxesSZ.count) * 100)
                Start-Sleep -Milliseconds 200
            }

            $endTime = Get-Date

            Write-host "Granting: $user to all Mailboxes completed on" $endTime -ForegroundColor Green

            <#
            test code
            $mailboxes = Get-Mailbox -ResultSize 10
            $user = "username@address.com"
            $counter = 0
            foreach ($mbx in $mailboxes){
                $counter++
                $display = $mbx.WindowsEmailAddress
                Add-MailboxPermission -Identity $mbx.UserPrincipalName -User $user -AccessRights fullaccess -AutoMapping:$false -InheritanceType all
                Add-RecipientPermission -Identity $mbx.UserPrincipalName -AccessRights SendAs -Trustee $user -Confirm:$false
                Write-Progress -Activity "Adding $user read & send Access to: $display" -Status "Processing $($counter) of $($mailboxes.count)" -CurrentOperation $mbx.WindowsEmailAddress -PercentComplete (($counter / $mailboxes.count) * 100)
                Start-Sleep -Milliseconds 200
            }
            #>  
}

# Removes username@address.com or other if specified access to all mailboxes in tenant
function RemovePermissionToAllMailboxes
{

    ConnectEXO

    $user  = Read-Host "Who do you want to remove all Read & Send Access to all Mailboxes? (Default is username@address.com if left empty)"
    
    if ($user)
            {
                $user = $user
            }

            else
            {
                $user = "username@address.com"
            }

            $startTime = Get-Date

            Write-Host "Make sure you run this function in a new PowerShell Tab and/or different computer if you want to do other work, as this will take a long time and lots of computer resources." -ForegroundColor Red

            Write-host "Removing: $user to all Mailboxes started on" $startTime -ForegroundColor Green

            Write-host "Gathering all Mailboxes through Display Name of A through F" -ForegroundColor Green

            $mailboxesAF = Get-Mailbox -ResultSize unlimited | ?{$_.DisplayName -match "^[A-F]"}
            
            $counter = 0
            foreach ($mbx in $mailboxesAF){
                $counter++
                $display = $mbx.WindowsEmailAddress
                Remove-MailboxPermission -Identity $mbx.UserPrincipalName -User $user -AccessRights fullaccess -InheritanceType all
                Remove-RecipientPermission -Identity $mbx.UserPrincipalName -AccessRights SendAs -Trustee $user -Confirm:$false
                Write-Progress -Activity "Removing $user read & send access to: $display" -Status "Processing $($counter) of $($mailboxesAF.count)" -CurrentOperation $mbx.WindowsEmailAddress -PercentComplete (($counter / $mailboxesAF.count) * 100)
                Start-Sleep -Milliseconds 200
            }


            Write-host "Gathering all Mailboxes through Display Name of G through L" -ForegroundColor Green

            $mailboxesGL = Get-Mailbox -ResultSize unlimited | ?{$_.DisplayName -match "^[G-L]"}
            
            $counter1 = 0
            foreach ($mbx in $mailboxesGL){
                $counter1++
                $display = $mbx.WindowsEmailAddress
                Remove-MailboxPermission -Identity $mbx.UserPrincipalName -User $user -AccessRights fullaccess -InheritanceType all
                Remove-RecipientPermission -Identity $mbx.UserPrincipalName -AccessRights SendAs -Trustee $user -Confirm:$false
                Write-Progress -Activity "Removing $user read & send access to: $display" -Status "Processing $($counter1) of $($mailboxesGL.count)" -CurrentOperation $mbx.WindowsEmailAddress -PercentComplete (($counter1 / $mailboxesGL.count) * 100)
                Start-Sleep -Milliseconds 200
            }


            Write-host "Gathering all Mailboxes through Display Name of M through R" -ForegroundColor Green

            $mailboxesMR = Get-Mailbox -ResultSize unlimited | ?{$_.DisplayName -match "^[M-R]"}
            
            $counter2 = 0
            foreach ($mbx in $mailboxesMR){
                $counter2++
                $display = $mbx.WindowsEmailAddress
                Remove-MailboxPermission -Identity $mbx.UserPrincipalName -User $user -AccessRights fullaccess -InheritanceType all
                Remove-RecipientPermission -Identity $mbx.UserPrincipalName -AccessRights SendAs -Trustee $user -Confirm:$false
                Write-Progress -Activity "Removing $user read & send access to: $display" -Status "Processing $($counter2) of $($mailboxesMR.count)" -CurrentOperation $mbx.WindowsEmailAddress -PercentComplete (($counter2 / $mailboxesMR.count) * 100)
                Start-Sleep -Milliseconds 200
            }


            Write-host "Gathering all Mailboxes through Display Name of S through Z" -ForegroundColor Green

            $mailboxesSZ = Get-Mailbox -ResultSize unlimited | ?{$_.DisplayName -match "^[S-Z]"}
            
            $counter3 = 0
            foreach ($mbx in $mailboxesSZ){
                $counter3++
                $display = $mbx.WindowsEmailAddress
                Remove-MailboxPermission -Identity $mbx.UserPrincipalName -User $user -AccessRights fullaccess -InheritanceType all
                Remove-RecipientPermission -Identity $mbx.UserPrincipalName -AccessRights SendAs -Trustee $user -Confirm:$false
                Write-Progress -Activity "Removing $user read & send access to: $display" -Status "Processing $($counter3) of $($mailboxesSZ.count)" -CurrentOperation $mbx.WindowsEmailAddress -PercentComplete (($counter3 / $mailboxesSZ.count) * 100)
                Start-Sleep -Milliseconds 200
            }


            $endTime = Get-Date

            Write-host "Removing: $user to all Mailboxes completed on" $endTime -ForegroundColor Green
}

# Removes username@address.com or other if specified access to a specific mailbox
function RemovePermissionToMailbox
{
    ConnectEXO

    $email  = Read-Host "Which Mailbox do you want to Remove Read & Send Access?"
    $user = Read-Host "Which do you want to remove from $email`? (Default is username@address.com if left empty)"
    
    if ($user)
            {
                $user = $user
            }

            else
            {
                $user = "username@address.com"
            }

            if((o365BoxExists($email)) -eq $true){
            Write-host "Removing: $user from $email" -ForegroundColor Red
            Remove-MailboxPermission -identity $email -user $user -AccessRights fullaccess -InheritanceType all -Confirm:$false 
            Remove-RecipientPermission $email -AccessRights SendAs -Trustee $user -Confirm:$false 
            }
            else
            {
            Write-Host "Mailbox `"$email`" does not exist!" -ForegroundColor red
            }
}

# Checks forwarding and gives the option to change it to something else or nothing
function CheckForwarding
{
    ConnectEXO

    $email = Read-Host "Which email do you want to check forwarding?"
    
    if((o365BoxExists($email)) -eq $true){
            
                $checkforwarding = Get-Mailbox $email | select ForwardingSmtpAddress

                $checkforwardingProperty = $checkforwarding.ForwardingSmtpAddress

                if(!$checkforwardingProperty)
                {
                    Write-Host "$email has no forwarding" -ForegroundColor Green
                    $optionCheck2 = Read-Host "Do you want to change it? <1> Yes <2> No? (<2> if left blank)"


                    switch($optionCheck2)
                {
                    1
                    {
                        $change = Read-Host "What email do you want to change it to?"
                        Set-Mailbox $email -ForwardingSmtpAddress $change
                        Write-Host "$email forwarding address has been changed to $change" -ForegroundColor Green
                    }

                    2
                    {
                        Return
                    }
                    
                    default
                    {
                        Return
                    }

                }



                }

                else
                {

                Write-Host "$email has $checkforwardingProperty"

                $optionCheck = Read-Host "Do you want to change it? <1> Yes <2> No? (<2> if left blank)"
                

                switch($optionCheck)
                {
                    1
                    {
                        $change = Read-Host "What email do you want to change it to?"
                        Set-Mailbox $email -ForwardingSmtpAddress $change
                        Write-Host "$email forwarding address has been changed to $change" -ForegroundColor Green
                    }

                    2
                    {
                        Return
                    }
                    
                    default
                    {
                        Return
                    }

                }
           }  

     }


     else 
     {
        Write-Host "$email doesn't exist" -ForegroundColor Red
        Return
     }

}

#############################################################
# Custom Add-ons Menu Items
# The create custom shortcuts
#############################################################

# Creates the menu in add-ons
function initMenu()
{
    $connectMenu = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Connect To...",$null,$null)

    $connectMenu.SubMenus.Add(
      "Connect to O365 + EXO/EOP + SPO + SFBO + CC + ARM + Exchange-OnPrem + Lync-OnPrem",
      {
            ConnectAll
      },
      "Control+Alt+A"
    )

    $connectMenu.SubMenus.Add(
      "Connect to Office 365 (O365)",
        {
	        ConnectO365
        },
      "Control+Alt+O"
    )

    $connectMenu.SubMenus.Add(
      "Connect to Exchange Online (EXO/EOP)",
      {
            ConnectEXO/EOP
      },
      "Control+Alt+Z"
    )

    $connectMenu.SubMenus.Add(
      "Connect to SharePoint Online (SPO)",
      {
            ConnectSPO	
      },
      "Control+Alt+S"
    )

    $connectMenu.SubMenus.Add(
      "Connect to Skype for Business Online (SFBO)",
      {
            ConnectSFBO	
      },
      "Control+Shift+S"
    )

    $connectMenu.SubMenus.Add(
      "Connect to Compliance Center (CC)",
      {
            ConnectCC	
      },
      "Control+Shift+C"
    )

    $connectMenu.SubMenus.Add(
      "Connect to Azure Rights Management (ARM)",
      {
            ConnectAzureRMS	
      },
      "Control+Alt+P"
    )

    $connectMenu.SubMenus.Add(
      "Connect to Exchange-OnPrem (Exchange-OnPrem)",
        {
            ConnectExchange        
        },
      "Control+Alt+X"
    )

    $connectMenu.SubMenus.Add(
      "Connect to Lync-OnPrem (Lync-OnPrem)",
        {
            ConnectLync       
        },
      "Control+Shift+M"
    )

    $connectMenu.SubMenus.Add(
        "Check Connection",
        {
            Get-PSSession | ft ComputerName, State, Availability -AutoSize
        },
        "Control+Alt+C"
    )

    $connectMenu.SubMenus.Add(
      "Disconnect All Remote Sessions Now (Timeouts after 15 mins of inactivity otherwise)",
        {
	        disconnectAll
        },
      "Control+Alt+D"
    )

    $openMenu = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Open...",$null,$null)

    $openMenu.SubMenus.Add(
       "Open Mailbox - (Chrome) as username@address.com or Other",
        {
            OpenChromeEmail 2>&1 | out-null
        },
        "Control+Alt+L"
    )

    $openMenu.SubMenus.Add(
       "Open Mailbox - (IE) as username@address.com or Other",
        {
            OpenIEEmail 2>&1 | out-null
        },
        "Control+Alt+K"
    )

    $openMenu.SubMenus.Add(
       "Open Mailbox - (Chrome) as username@address.com",
        {
            OpenChromeAdmin 2>&1 | out-null
        },
        "Control+Alt+E"
    )

    $openMenu.SubMenus.Add(
       "Open Mailbox - (IE) as username@address.com",
        {
            OpenIEAdmin 2>&1 | out-null
        },
        "Control+Alt+I"
    )

    $adminMenu = $psISE.CurrentPowerShellTab.AddOnsMenu.SubMenus.Add("Admin...",$null,$null)

    $adminMenu.SubMenus.Add(
       "Grant username@address.com or Other Read & Send Access to all Mailboxes",
        {
            GrantPermissionToAllMailboxes
        },
        "Control+Shift+K"
    )

    $adminMenu.SubMenus.Add(
       "Remove username@address.com or Other Read & Send Access to all Mailboxes",
        {
            RemovePermissionoAllMailboxes
        },
        "Control+Shift+A"
    )

    $adminMenu.SubMenus.Add(
       "Remove Read & Send Access to Mailbox",
        {
            RemovePermissionToMailbox
        },
        "Control+Shift+E"
    )

    $adminMenu.SubMenus.Add(
       "Get User Info",
        {
            Get-UserInfo
        },
        "Control+Alt+G"
    )

    $adminMenu.SubMenus.Add(
       "Check Forwarding",
        {
            CheckForwarding
        },
        "Control+Shift+Q"
    )

}

initMenu 2>&1 | out-null