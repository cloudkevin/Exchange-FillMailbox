Param(
    [Parameter(Mandatory=$True)]
        [string]$Server,
    [Parameter(Mandatory=$True)]
        [string]$TargetMailbox,
    [Parameter(Mandatory=$False)]
        [int]$NumDaysBack = 120,
    [Parameter(Mandatory=$False)]
        [int]$MsgsPerDay = 5,
    [Parameter(Mandatory=$False)]
        [System.Int64]$MsgSize = 100kb
)

$userCreds = Get-Credential -Message "Enter Credentials for Account with Impersonation Role..."

Function ConnectExchangeService($Server)
{
[Reflection.Assembly]::LoadFile("C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll") | Out-Null
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$service.Url = new-object Uri("https://$Server/ews/exchange.asmx")
return $service
}
$service = ConnectExchangeService $Server
$service.Credentials = New-Object System.Net.NetworkCredential($userCreds.UserName, $userCreds.Password)
# $TargetMailbox = "target@domain.com"
$service
## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
  
## Code From http://poshcode.org/624
## Create a compilation environment
$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
$Compiler=$Provider.CreateCompiler()
$Params=New-Object System.CodeDom.Compiler.CompilerParameters
$Params.GenerateExecutable=$False
$Params.GenerateInMemory=$True
$Params.IncludeDebugInformation=$False
$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
$TAAssembly=$TAResults.CompiledAssembly

## We now create an instance of the TrustAll and attach it to the ServicePointManager
$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

## end code from http://poshcode.org/624
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$TargetMailbox)     
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)    

$Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailbox)

$PR_MESSAGE_DELIVERY_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E06, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
$PR_CLIENT_SUBMIT_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0039, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
$PR_Flags = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3591, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  

$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$TargetMailbox)     
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)    

$Attachment = New-Object System.IO.FileStream $Env:Temp\test.txt, Create, ReadWrite
$Attachment.SetLength($MsgSize)
$Attachment.Close()

[datetime]$StartDate = Get-Date

for($i=0; $i -lt $NumDaysBack; $i++) {
    for($j=0; $j -lt $MsgsPerDay; $j++) {
        $Message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($Service)
        $Message.Subject = "Date / Time Test Day #" + ($i+1)
        $Message.Body = "This is a test message used to test date operations on a mailbox."
        $Message.From = $TargetMailbox
        $Message.ToRecipients.Add($TargetMailbox) | Out-Null
        $Message.SetExtendedProperty($PR_Flags,"1")
        $Message.SetExtendedProperty($PR_MESSAGE_DELIVERY_TIME, $StartDate.AddDays($i * -1))
        $Message.SetExtendedProperty($PR_CLIENT_SUBMIT_TIME, $StartDate.AddDays($i * -1))
        $Message.Attachments.AddFileAttachment($attachment.Name) | Out-Null
        $Message.Save($Inbox.Id)
    }
    Write-Progress -activity "Creating Messages..." -status "Day $i of $NumDaysBack" -percentcomplete (($i/$NumDaysBack)*100)
}