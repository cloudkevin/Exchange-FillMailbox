# Inputs
-Server = your mail server  
-TargetMailbox = mailbox to fill with mail  
-NumDaysBack = number of days to backfill with mail  
-MsgsPerDay = number of messages to generate for each day  
-MsgSize = size of each test message  

# Example
```./fillMailbox.ps1 -Server yourserver.local -Target user@domain.com -NumDaysBack 365 -MsgsPerDay 20 -MsgSize 200kb```

# Requirements
You must have an account with application impersonation rights configured  
```New-ManagementRoleAssignment -name:impersonationAssignmentName -Role:ApplicationImpersonation -User:serviceAccount```
