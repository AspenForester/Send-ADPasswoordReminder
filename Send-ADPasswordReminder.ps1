<#
.SYNOPSIS
    Enumerates enabled users, and sends a reminder to change the password.
.DESCRIPTION
    Enumerates enabled users, and sends an email reminder to change the password. 
    Multiple reminders are handled by the array "days" containing numbers of days prior to expiration.
.EXAMPLE
    PS C:\> Send-ADPasswordReminder 
.EXAMPLE
    PS C:\> $SearchRoot = "ou=My Data Center,ou=users,ou=na,dc=foo,dc=local"
    PS C:\> Send-ADPasswordReminder -SearchRoot $SearchRoot
  .PARAMETER SearchBase
    Optional parameter to specify the search base for the AD lookup
  .PARAMETER Days
    Edit this array of values, adding a value for each reminder you want sent 
    equal to the number of days before the user's password is set to expire.
  .PARAMETER SMTPServer

  .PARAMETER From

  .PARAMETER LogPath

  .PARAMETER MaxLogLines
    The maximum number of lines that the log file is allowed to grow to.
    Older lines will be truncated
  .INPUTS
    Inputs to this cmdlet (if any)
  .OUTPUTS
    Output from this cmdlet (if any)
  .NOTES
    Written By: JB Lewis
    Originally written: 27-July-2011 
    Rewritten: 11/01/2017
  .COMPONENT
    The component this cmdlet belongs to
  .ROLE
    The role this cmdlet belongs to
  .FUNCTIONALITY
    The functionality that best describes this cmdlet
#>
function Send-ADPasswordReminder
{
    [CmdletBinding(SupportsShouldProcess)]
    param(
        # SearchBase for AD Query
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true)]
        [string]
        $SearchBase,

        # SMTP Server 
        [Parameter(Mandatory = $true)]
        [String]
        $SMTPServer,

        # From address for the email
        [Parameter(Mandatory = $true)]
        [String]
        $From,

        # Days to compare for password expiry notification
        [Parameter(Mandatory = $false)]
        [int[]]
        $Days = (14, 3, 1),

        # Log File location and path
        [Parameter(Mandatory = $false)]
        [String]
        $LogPath = (Split-Path -parent $MyInvocation.MyCommand.Definition),

        # Maximum Logfile length
        [Parameter(Mandatory = $false)]
        [int]
        $MaxLogLines = 5000
    )
    begin
    {
        # Input handling
        if (-not (get-module -ListAvailable -Name ActiveDirectory))
        {
            Throw "This Function requires the ActiveDirectory Module"
        }
        if ($from -notmatch "^[a-zA-Z0-9.!#$%&’*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$")
        {
            Throw "$From is not a properly formed email address!"
        }

        function Resize-log($LogFilePath, $LinesToKeep)
        {
            $f = Get-content $LogFilePath
            $count = $f.count
            if ($count -gt $LinesToKeep)
            {
                $f | select -last $LinesToKeep | Set-content $LogFilePath
            }
        }
    }
    process
    {
        $SearchRoot = "thcg.net/na/users/plymouth data center"
        # Use the following for Get-ADUser
        # $SearchRoot = "ou=My Data Center,ou=users,ou=na,dc=foo,dc=local"

        $LogFile = "$LogPath\$(($MyInvocation.MyCommand).Name).log"

        $mailer = new-object Net.Mail.SmtpClient ($smtpserver)
        $from = $from

        $now = Get-Date
        Add-Content -Path $LogFile -Value "$(Get-Date) - Starting"

        # The following provides the same results as the Get-QADUser cmdlet
        $users = Get-ADUser -Filter {enabled -eq  $True} `
            -Properties "msds-UserPasswordExpiryTimeComputed", "mail" `
            -SearchBase $SearchRoot `
            | Select GivenName, Surname, mail, @{Name = 'PasswordExpires'; Expression = {[DateTime]::FromFileTime($_."msds-UserPasswordExpiryTimeComputed")}
        }
	
        Foreach ($user in $users)
        {
            $when = (New-TimeSpan -Start $now -End $user.PasswordExpires).Days
            If ($when -in $Days)
            {
                if ($When -gt 1) {$plural = "s"} else { $plural = ""}
                # Send an email
                $msg = new-object Net.Mail.MailMessage
                $msg.From = $from
                $msg.To.Add($($user.Email))
                $msg.subject = "Your Password expires in $Day day$plural"
                $msg.Body = @"
<html>
<head>
	<style type="text/css">
		body {
			font-family: Verdana, Arial, Helvetica, sans-serif;
			font-size: x-small;
			background: #fff;
			color: #000;
			}
		table {
			border-width: 1px;
			border-style: solid;
			border-color: #003878;
			border-spacing: 15px;
			background-color: white;
			}
	</style>
</head>
<body>
	<table width="620">
		<tr>
			<td align="left" valign="top">
				<p><strong>Dear $($user.FirstName) $($user.LastName),</strong></p>
				<p>Just a friendly reminder that your Domain password expires $Day day$plural from Now.</p>
				<ul>
				<li>
				<P>Option 1</P>
				<P>Please update your Domain password by pressing the Ctrl + Alt + Delete keys simultaneously from your Windows workstation and selecting 'Change Password' from the available options.</P>
				<P>Note: If you are a remote user, you will have to be connected to the VPN for this to work. Please restart your workstation after changing your password.</P>
				</li>
				<li>
				<P>Option 2</P>
				<P>Connect to [some hyperlink] and login.  After you login you will be given the option to Change Password.  After changing your password, you should then be able to connect VPN by logging into [Some other link].  After you log into VPN please hit Ctrl + Alt + Delete key simultaneously and then “Lock” your computer and “Unlock” your computer with your new password.</P>
				<P>Note: This option can also be used if you miss the date that your password expires.</P>
				</li>
				</ul>
				<P><strong>For iPhone and iPad users:</strong>  After you have completed the steps listed above, follow these next steps to apply your new password to your iPad and/or iPhone.  Open the <i>Settings</i> icon from the home screen of your device, open <i>Mail, Contacts, Calendars</i>, open the account called <i>(Your company) Exchange</i>, then open the account listed with your company email address, finally enter your new password in the <i>Password</i> field.</P>
				<P>Remember, if you do not change your password before it expires you could be locked out from accessing internal company resources until an Administrator unlocks your account.</p>
				<P><i>Passwords must be a minimum of 8 characters and Require 3 of the 4 characteristics below:</i></p>
				<ul>
				<li>Lower case letter</li>
				<li>Capital Letter</li>
				<li>Numeral</li>
				<li>Special Characters ($,#,*, !, etc.)</li>
				</ul>
				<P>*Note: Additional requirements will limit password re-use, frequent password changes, as well as using your first or last name.</p>
				<p>If you have any questions or need further assistance, please contact your IT Service Desk.</p>
				<p>&nbsp;</p>
				<p>Thank you,</p>
				<p><strong>YOur Company - IT Service Desk </strong></p>
				<p><strong>808-555-2345</strong></p>
			</td>
		</tr>
	</table>
</body>
</html>
"@
                $msg.IsBodyHTML = $true 
			
                if ($pscmdlet.ShouldProcess("$($user.GivenName) $($user.Surname)", "Notify ")) 
                {
                    $mailer.send($msg)
                }
                $msg = $null
                # Write-Host "$user's password will expire in $day days, on $($user.PasswordExpires)"
                # It would be easy to add some logging here.
                try 
                {
                    Add-Content -Path $LogFile -Value "$(Get-Date) - $user notified $day days prior to expiration" -ErrorAction Stop
                }
                Catch 
                {
                    Write-Warning "Unable to write to log $LogFile"
                }   
            }
        }
    }
    end
    {
        If ($MaxLogLines -gt 0)
        {
            try 
            {
                Resize-Log $LogFile, $MaxLogLines 
            }
            catch 
            {
                Write-Warning "Unable to resize $LogFile!"
            }
        }
    }
}
