
# ensure you are in the directory that contains:
# * this file, 
# * the Provision-AzureAD.ps1 file, 
# * the IDE01.CSV file that is to be imported.
cd Z:\U\SKY\D\NOSYNC\Repositories\sct-provision-ad

# Select and run this method to ask you for your Azure AD account name and password
if (!$credentials) {
  $credentials = Get-Credential
}

# prove that the above account is an Azure AD Identity (not just a Microsoft Account!):
Connect-MsolService -Credential $credentials


#Load the file that contains the method:
.\Provision-AzureAD.ps1

# Invoke the Command:
Provision-AzureAD -ideFileName:'IDE01.csv' -ideRoleType:'Student' -forceKeepPasswordYears:'Year1,Group00' -adDomainName:'powershelltesting2.onmicrosoft.com' -newADUserDefaultPassword:'W3lcome!' -adPrincipalCredential:$credentials -Verbose

                        