# Get-MailboxFolderPermissionEWS
## Description

The purpose of the script is to get for each folder granted permissions for given mailboxes.

## Requirements

# Permissions: Either FullAccess or ApplicationImpersonation

# [Exchange Web Services Managed API](http://go.microsoft.com/fwlink/?LinkId=255472)


## Examples

#run the script against a single mailbox
```
.\Get-MailboxFolderPermissionEWS.PS1 -EmailAddress trick@adatum.com
```

#pipe all the mailboxes on database USEX01-DB01 into the script
```
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1
```

#pipe all the mailboxes on database USEX01-DB01 into the script and use 10 threads
```
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1 -Threads 10 -MultiThread
```

#pipe all the mailboxes on database USEX01-DB01 into the script and use 10 threads
```
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1 -Threads 10 -MultiThread
```

#same as the previous, but this time using different credentials and impersonate
```
$SA= Get-Credential adatum\serviceaccount
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1 -Threads 10 -MultiThread -Server ex01.adatum.com -UseDefaultCred 0 -Impersonate -Credentials $SA
```

#to export the result just use the Export-CSV CmdLet
```
Get-Mailbox -Database USEX01-DB01 -ResultSize unlimited | .\Get-MailboxFolderPermissionEWS.PS1 | Export-Csv -NoTypeInformation -Path .\result.csv
```

#scan multiple mailboxes for invalid permission entries using ApplicationImpersonation, give credentials and multithreading
```
Get-MailboxDatabase dee16dag01* | Get-Mailbox -ResultSize Unlimited | .\Get-MailboxFolderPermissionEWS.PS1 -Impersonate -Credentials $cred -MultiThread -SearchUnknownOnly
```

## Required Parameters

### -EmailAddress

The e-mail address of the mailbox, which will be checked. The script accepts piped objects from Get-Mailbox or Get-Recipient.

## Optional Parameters

### -Credentials

Credentials you want to use. If omitted current user context will be used.

### -Impersonate

Use this switch, when you want to impersonate.

### -CalendarOnly

By default the script will enumerate all folders. If you want to limit to folders with type "IPF.Appointment" use this switch.

### -RootFolder

From where you want to start the search for folders. Default is "MsgFolderRoot". Other posible value is "Root".

### -Server

By default the script tries to retrieve the EWS endpoint via Autodiscover. If you want to run the script against a specific server, just provide the name in this parameter. Not the URL!

### -Threads

How many threads will be created. Be careful as this is really CPU intensive! By default 10 threads will be created. I limit the number of threads to 20.

### -MultiThread

If you want to run the script multi-threaded use this switch. By default the script don't use threads.

### -MaxResultTime

The timeout for a job, when using multi-threads. Default is 240 seconds.

### -MrMapi

The path to MrMapi.exe. Only full path is accepted!

### -UseMrMapi

Switch to tell the script to use MrMapi or not.

### -TrustAnySSL

Switch to trust any certificate.

### -SearchUnknownOnly

Returns only folders with invalid permissions.

## Links

# [Exchange Online migration and TooManyBadItemsPermanetException](https://ingogegenwarth.wordpress.com/2018/06/15/exo-mig-invalidperm/)

# [The good, the bad and sIDHistory](https://ingogegenwarth.wordpress.com/2015/04/01/the-good-the-bad-and-sidhistory/)

# [Get mailbox folder permissions using EWS multithreading](https://ingogegenwarth.wordpress.com/2015/04/16/get-mailbox-folder-permissions-using-ews-multithreading/)
