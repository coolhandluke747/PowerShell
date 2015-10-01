# These are the commands I used on Office 365 to stop a on-going phishing attack happening currently on your tenant.

# First get a .CSV off all your mailboxes.
Get-Mailbox -ResultSize Unlimited | select PrimarySmtpAddress | Export-Csv PrimarySmtpAddress.csv

# Then import this .CSV into this next command and run a search on each mailbox to delete the phishing email.
Import-Csv "PrimarySmtpAddress.csv" | Foreach {Search-Mailbox -Identity $_.PrimarySmtpAddress -SearchQuery 'From:email@email.com and Sent:08/25/2014' -DeleteContent -Force}?
