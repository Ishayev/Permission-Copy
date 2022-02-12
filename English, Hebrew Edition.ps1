Import-Module exchangeonlinemanagement
Connect-ExchangeOnline
Connect-AzureAD
$List=Get-Mailbox | select WindowsEmailAddress

function Show-Menu
{
    param ([string]$Title = 'Daniel Ishayev 365 Menu')
    Clear-Host
    Write-Host "================ $Title ================"
    Write-Host "1: Get Contacts Permissions."
    Write-Host "2: Get Calendar Permissions."
    Write-Host "3: Remove Contacts Permissions English(After 1 ONLY)."
    Write-Host "4: Remove Calendar Permissions(After 2 ONLY)."
    Write-Host "5: Add Contacts Permissions(After 1 ONLY)."
    Write-Host "6: Add Calendar Permissions(After 2 ONLY)."

    Write-Host "Q: Press 'Q' to quit."
}
do
 {
  
   Show-Menu
     $selection = Read-Host "Please make a selection:"
     switch ($selection)
     {
     '1' {

     $365Permissions = New-Object PSObject
     $365Permissions= @() 

$user=read-host "Please Enter Display Name of the User"
for($i=0;$i -lt $List.count;$i++)
{
$E=Get-AzureADUserLicenseDetail -ObjectId $list.WindowsEmailAddress[$i] | select SkuPartNumber
## 'not empty' 
if ($E) {
$test=':\contacts'
$test1=':\אנשי קשר'
$g=$list[$i].WindowsEmailAddress+$test
$g1=$list[$i].WindowsEmailAddress+$test1
$b=Get-MailboxFolderPermission $g
$b1=Get-MailboxFolderPermission $g1
$d=$b.user.displayname
$d1=$b1.user.displayname
if($d -eq $user)
{
  $365Permissions+= @{EmailEnglish=$list[$i].WindowsEmailAddress}

}
if($d1 -eq $user)
{
  $365Permissions+= @{EmailHebrew=$list[$i].WindowsEmailAddress}

}
}
}

$checking=$365Permissions.emailenglish
$checking1=$365Permissions.emailhebrew
$checking | out-file -filepath c:\Eng.txt
$checking1 | out-file -filepath c:\Heb.txt
C:\Heb.txt
C:\Eng.txt

}

     '2' {
     $365Permissions = New-Object PSObject
     $365Permissions= @() 

$user=read-host "Please Enter Display Name of the User"
for($i=0;$i -lt $List.count;$i++)
{
$E=Get-AzureADUserLicenseDetail -ObjectId $list.WindowsEmailAddress[$i] | select SkuPartNumber
## 'not empty' 
if ($E) {
$test=':\calendar'
$test1=':\לוח שנה'
$g=$list[$i].WindowsEmailAddress+$test
$g1=$list[$i].WindowsEmailAddress+$test1
$b=Get-MailboxFolderPermission $g
$b1=Get-MailboxFolderPermission $g1
$d=$b.user.displayname
$d1=$b1.user.displayname
if($d -eq $user)
{
  $365Permissions+= @{EmailEnglish=$list[$i].WindowsEmailAddress}

}
if($d1 -eq $user)
{
  $365Permissions+= @{EmailHebrew=$list[$i].WindowsEmailAddress}

}
}
}
$checking=$365Permissions.emailenglish
$checking1=$365Permissions.emailhebrew
$checking | out-file -filepath c:\Eng.txt
$checking1 | out-file -filepath c:\Heb.txt
C:\Heb.txt
C:\Eng.txt
}


     '3' {
          $user=read-host "Please Enter Display Name of the User"
        for($i=0;$i -lt $365Permissions.count;$i++)
        {
          $r=$checking[$i]+$test
          $r1=$checking1[$i]+$test1
          Remove-MailboxFolderPermission -Identity $r -User $user
          Remove-MailboxFolderPermission -Identity $r1 -User $user
          }
       }
         
     '4' {
          $user=read-host "Please Enter Display Name of the User"
        for($i=0;$i -lt $365Permissions.count;$i++)
        {
          $r=$checking[$i]+$test
          $r1=$checking1[$i]+$test1
          Remove-MailboxFolderPermission -Identity $r -User $user
          Remove-MailboxFolderPermission -Identity $r1 -User $user
          }
       }
     '5' {
        $user1=read-host "Please Enter Display Name of the User"
        for($i=0;$i -lt $365Permissions.count;$i++)
        {
          $r=$checking[$i]+$test
          $r1=$checking1[$i]+$test1
          Add-MailboxFolderPermission -Identity $r -User $user1 -AccessRights Editor
          Add-MailboxFolderPermission -Identity $r1 -User $user1 -AccessRights Editor
          
}
}
     '6' {
         $user1=read-host "Please Enter Display Name of the User"
        for($i=0;$i -lt $365Permissions.count;$i++)
        {
          $r=$checking[$i]+$test
          $r1=$checking1[$i]+$test1
          Add-MailboxFolderPermission -Identity $r -User $user1 -AccessRights Editor
          Add-MailboxFolderPermission -Identity $r1 -User $user1 -AccessRights Editor
}
}
}
}
 until ($selection -eq 'q')

         
