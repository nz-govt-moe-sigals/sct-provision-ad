# May run into:forceChangedPassword
# https://support.microsoft.com/en-nz/help/2669552/-the-term-cmdlet-name-is-not-recognized-error-when-you-try-to-run-azure-active-directory-for-windows-powershell-cmdlets
# which will get you to download AdministrationConfig-V1.1.166.0-GA.msi
# http://connect.microsoft.com/site1164/Downloads/DownloadDetails.aspx?DownloadID=59185

function Provision-AzureAD {
  <#
  .SYNOPSIS
  An Idempotent script to take a NZ Student Management System (SMS)'s IDE file and provision Azure AD Accounts.
  .DESCRIPTION
  The script imports a standard IDE file exported from the SMS, and if the required data (First, Last, Email) is filled in, ensures that an Azure Account exists.
  Intended to be run by SIs when first setting up a school for Azure AD, and subsequently by school administrators as needed when students come and go.
  The script is a baseline one, that System Integrators (SIs) can extend to suite their individual needs.
  .EXAMPLE
  Provision-AzureAD
  .EXAMPLE
  Provision-AzureAD -adDomainName:myschool.school.nz
  .EXAMPLE
  Provision-AzureAD -adDomainName:myschool.school.nz -newADDefaultPassword:W3lcome!
  .EXAMPLE
  Provision-AzureAD -adDomainName:myschool.school.nz -newADDefaultPassword:W3lcome! -adPrincipalName:admin@myschool.school.nz -newADDefaultPassword:W3lcome!
  .PARAMETER adDomainName
  The DNS domain name associated to the school's Azure AD tenancy (eg: 'someschool.school.nz'). The default value is null.
  .PARAMETER adPrincipalName
  The Azure AD Tenancy admin's principal name (eg: 'admin3@someschool.school.nz'). The default value is null.
  .PARAMETER adPrincipalSecurePassword
  The Azure AD Tenancy admin's principal name (eg: 'V3rySecure!'). The default value is null.
  .PARAMETER adPrincipalCredential
  Instead of passing name and password, the script can be invoked with a Credential created using the Get-Credential command. The default value is null.
  .PARAMETER ideRoleType
  A comma separated list of RoleTypes to import ('Student,'TeachingStaff', etc.). The default value is 'Student'
  .PARAMETER forceKeepPasswordYears
  By default, new users are created with passwords they must change when they first log in. For younger primary schools, Passwords can be created such that young students are not required to change their password. This is determined by their belonging to one of a comma separated list of the Home Groups (eg: 'Year1,Group00')
  .PARAMETER ideFileName
  The path to the IDE file to import. By default it searches in the script's startup folder for 'IDE01.csv'. The default value is './IDE01.csv'
  .PARAMETER ideFilePasswordColumnName
  The name of the optional column in the imported IDE file that contains the password to use if creating a new user. The default value is 'Password'
  .PARAMETER newADUserDefaultPassword
  If no ideFilePasswordColumnName found, falls back to using this value when creating a new user. The default value is null (a random value will be generated).
  #>
   [CmdletBinding()]
   Param (
     [Parameter(mandatory=$false)] [string]$adDomainName, 
     [Parameter(mandatory=$false)] [string] $adPrincipalName,     
     [Parameter(mandatory=$false)] [SecureString] $adPrincipalSecurePassword, 
     [Parameter(mandatory=$false)] [System.Management.Automation.PSCredential] $adPrincipalCredential,
     [Parameter(mandatory=$false)] [string] $ideFileName = "./IDE01.csv", 
     [Parameter(mandatory=$false)] $ideRoleType="Student", 
     [Parameter(mandatory=$false)] $forceKeepPasswordYears=$null,
     [Parameter(mandatory=$false)] $displayNameSeparator=" ",
     [Parameter(mandatory=$false)] [string]$ideFilePasswordColumnName = "Password",
     [Parameter(mandatory=$false)] [string]$newADUserDefaultPassword=$null,
     [Parameter(mandatory=$false)] [string]$reportColumnSeparator="`t"
     )

  # need an array in which to record failed provisioning,
  # which is what will be returned.
  [PSObject[]]$successRecords = @();
  [PSObject[]]$failedRecords = @();
  [PSObject[]]$tmpArray = @();
 
  [int]$recordsFoundCount = 0;
  [int]$recordsProcessedCount = 0;
  [int]$recordsNotReadyCount = 0;
  [int]$recordsAlreadyExistingCount = 0;
  [int]$recordsFailedCount = 0;
  [int]$recordsCreatedCount = 0;

  [PSObject]$tmp = $null;
  [PSObject[]]$reportItems = @();
  [PSObject[]]$ideRecords = $null;
  [PSObject]$adUser= $null;

  Write-Verbose -Message "This works if the following prerequisites are met:";
  Write-Verbose -Message "* The organisation has an AD Tenant (eg: someorg.onmicrosoft.com)";
  Write-Verbose -Message "* The organisation's AD Tenant has been associated to a DNS record (eg: someorg.com)";
  Write-Verbose -Message "* The organisation's AD Tenant's DNS record has been verified (or you'll get login errors).";
  Write-Verbose -Message "* The organisation's AD Tenant's DNS record has been verified.";
  Write-Verbose -Message "* The organisation's AD Tenant's DNS record is (optionally) set as the default DNS domain name.";
  Write-Verbose -Message "* The account of the user invoking this tool is a Work or Business AD Account (Microsoft Account won't do).";
  Write-Verbose -Message "* The user invoking this method has been invited into the above tenant as an AD (limited) User Admin.";
  Write-Verbose -Message "In other words:";
  Write-Verbose -Message "James Dean (jd@someITShop.com) -- a consultant System Integrator invited by a school or org";
  Write-Verbose -Message "to perform this operation on their behalf -- must be first invited into the schools";
  Write-Verbose -Message "Tenancy, and given enough rights to create Users, ie: AD (limited) User Admin.";
  Write-Verbose -Message "";
  Write-Verbose -Message "Tests:";
  Write-Verbose -Message "* Ensure jd@someITshop.com can log in to portals.azure.com";
  Write-Verbose -Message "* The following two commands, typed in Powershell's IDE don't give an error:";
  Write-Verbose -Message '  $credentials = Get-Credential';
  Write-Verbose -Message '  Connect-MsolService -Credential $credentials';
  Write-Verbose -Message "The above can commonly can fail if";
  Write-Verbose -Message "* The user is using a Microsoft Acount.";
  Write-Verbose -Message "* AD PS tools are not installed: http://connect.microsoft.com/site1164/Downloads/DownloadDetails.aspx?DownloadID=5918";  
  Write-Verbose -Message "If the above work it means that James Dean's Name and Password are accepted.";
  Write-Verbose -Message "";
  Write-Verbose -Message "--------------------------------------------------";
  Write-Verbose -Message "Starting....";
  

  
  $Error.Clear();



  # ensure that variables are in expected type.

  # Ensure credentials:
  if ($adPrincipalCredential -eq $null) {
    if ( ([string]::IsNullOrEmpty($adPrincipalName)) -or ($adPrincipalSecurePassword.length -eq 0) ){
      Write-Host -Message "Missing 'adPrincipalName' or 'adPrincipalSecurePassword' so showing dialog to requests Credentials needed to create 'adPrincipalCredential'.";
      $adPrincipalCredential = Get-Credential
      if ($Error.Count -gt 0){
        Write-Host -Message "Error creating Credential:"
        Write-Host -Message $Error[0].Exception.Message;
        Write-Host "Tip: Consider using the '-Verbose' flag to provide insight as to usage and execution". 
        throw "Error creating Credential.";
      }
    }else{
      #$adPrincipalSecurePassword  = ConvertTo-SecureString $adPrincipalPassword -AsPlainText -Force
      Write-Verbose -Message "Creating 'adPrincipalCredential' from provided '-adPrincipalName' and '-adPrincipalSecurePassword'.";
      $adPrincipalCredential = New-Object System.Management.Automation.PSCredential ($adPrincipalName, $adPrincipalSecurePassword);
      if ($Error.Count -gt 0){
        Write-Host -Message "Error creating Credential from privded '-adPrincipalName' and '-adPrincipalSecurePassword':";
        Write-Host -Message $Error[0].Exception.Message;
        Write-Host "Tip: Consider using the '-Verbose' flag to provide insight as to usage and execution". 
        throw "Error creating Credential.";
      }#~if
    }#~else
  }#~if

  if (($adPrincipalCredential.UserName -eq $null) -or ($adPrincipalCredential.Password -eq $null)){
    Write-Verbose -Message "Error creating Credential: '-adPrincipalCredential' is incomplete.";
    Write-Host "Tip: Consider using the '-Verbose' flag to provide insight as to usage and execution". 
    throw "Error creating Credential.";
  }

  # Ensure the parameter used to define which years should be ok to leave as fixed passwords (eg, early years) is an array:
  if ($forceKeepPasswordYears -eq $null){
    Write-Verbose -Message "'-forceKeepPasswordYears' not provided. That's ok.";
    $forceKeepPasswordYears = @();
  }
  else{
    if ($forceKeepPasswordYears -is [string] ){
      Write-Verbose -Message "'forceKeepPasswordYears' provided. Splitting it into an array...";
      $forceKeepPasswordYears = ($forceKeepPasswordYears -as [string]).Split([System.StringSplitOptions]::RemoveEmptyEntries);
    }
    else {
      throw "Cannot determine Type of 'forceKeepPasswordYears' parameter. Should be an array or CSV string.";
    }
  }
  $forceKeepPasswordYears = $forceKeepPasswordYears.ForEach({$_.Trim()});


  # Ensure parameter used to define which roles to import, is an array:
  if ($ideRoleType -eq $null){ 
    Write-Verbose -Message "'-ideRoleType' not provided. That's ok.";
    $ideRoleType = @();
  }
  else{
    if ($ideRoleType -is [string]){
      Write-Verbose -Message "'-ideRoleType' provided. Splitting it into an array...";
      $ideRoleType = ($ideRoleType -as [string]).Split([System.StringSplitOptions]:: RemoveEmptyEntries);
    }
    else {
      throw "Cannot determine Type of '-ideRoleType'. Should be either an array or CSV string."
    }
  }
  $ideRoleType = $ideRoleType.ForEach({$_.Trim()});


  #Ensure we have a domain Name
  if ([string]::IsNullOrEmpty($adDomainName)){
    Write-Host -Message "'adDomainName' was not provided, so have to ask what it is...";
    $adDomainName = Read-Host -Prompt "Enter the Domain Name (eg: 'company.com')";
    if ([string]::IsNullOrEmpty($adDomainName)){
      throw "An AD Domain Name (ie: 'adDomainName') is required";
    }
  }

  #Ensure we have char used to Build DisplayName
  if ($displayNameSeparator -eq $null){
    Write-Display -Message "'-displayNameSeparator' not provided. That's ok. So setting it to ' ' ...";
    $displayNameSeparator = " ";
  }

  if (!$reportColumnSeparator) {
    Write-Host -Message "'-reportColumnSeparator' not provided. That's ok. So setting it to '\t' ...";
    $reportColumnSeparator = "\t"
  }

  #Ensure we have the CSV to import:
  if ([string]::IsNullOrEmpty($ideFileName)){
    Write-Verbose -Message "'-ideFileName' not provided. That's ok. So setting it to './IDE01.csv' (ie, in current directory) ...";
    $ideFileName = "./IDE01.csv";
  }

  # Ensure we are importing Students
  if ([string]::IsNullOrEmpty($ideRoleType)){
    Write-Verbose -Message "'-ideRoleType' not provided. That's ok. So setting it to 'Student' (ie, in current directory) ...";
    $ideRoleType = "Student";
  }


  if ([string]::IsNullOrEmpty($newADUserDefaultPassword)){
    Write-Verbose -Message "'-newADUserDefaultPassword' not provided. That's ok. If CSV import has a 'Password' column, will use that, otherwise, will generate from scratch...";
  }

  # Get filtered records, and weed out ones that can't be processed:
  Write-Verbose -Message "Importing CSV from CSV...";
  [PSObject[]] $ideRecords = Import-CSV -Path $ideFileName | Where-Object { ($ideRoleType -as [array]).Contains($_.mlepRole)};
  if ($Error.Count -gt 0){
    Write-Verbose -Message "Error reading File.";
    Write-Verbose -Message $Error[0].Exception.Message;
    return;
  }

  $recordsFoundCount = $ideRecords.Count;
  Write-Verbose -Message "...csv imported.";

  # If we have a valid Credential, should be able to sign in to AD:

  # Connect to Azure AD (via Office 365 API):
  Write-Verbose -Message "...Connecting to Azure AD (via O365 APIs) with 'adPrincipalCredential'...";
  $Error.Clear();
  Connect-MsolService -Credential $adPrincipalCredential;
  if ($Error.Count -gt 0){
    Write-Host -Message $Error[0].Exception.Message;
    throw "Could not connect to Azure AD with provided Credentials.";
  }
  
  # Tried Jobs. Ran *slower* than a plain iteration:
  Write-Verbose -Message "Beginning to iterate through records in imported IDE file.";
  foreach ($ideRecord in $ideRecords) {
    Write-Verbose -Message "Processing Record...";

    # embed date for final report 
    $tmp = Get-Date -Format:"yyyyMMdd-HHmmss";
    Add-Member -InputObject $ideRecord -MemberType NoteProperty -Name "DateProcessed" -Value $tmp;
    Add-Member  -InputObject $ideRecord -MemberType NoteProperty -Name "DisplayName" -Value $null;
    if(!(Get-Member -inputobject $ideRecord -name "Password" -Membertype Properties)){
      Add-Member -InputObject $ideRecord -MemberType NoteProperty -Name "Password" -Value $null;
    }
    Add-Member -InputObject $ideRecord -MemberType NoteProperty -Name "ForceKeepPassword" -Value $false;

    Add-Member -InputObject $ideRecord -MemberType NoteProperty -Name "Success" -Value $false;
    Add-Member -InputObject $ideRecord -MemberType NoteProperty -Name "Comment" -Value  $null;

    if ([string]::IsNullOrEmpty($ideRecord.mlepFirstName)){
      Write-Verbose -Message "Missing mlepFirstName in CSV record.";
      $ideRecord.Comment = "Missing FirstName.";
      $recordsNotReadyCount += 1;
      $failedRecords += $ideRecord;
      continue;
    }
    if ([string]::IsNullOrEmpty($ideRecord.mlepLastName)){
      Write-Verbose -Message "Missing mlepLastName in CSV record.";
      $ideRecord.Comment = "Missing FirstName.";
      $recordsNotReadyCount += 1;
      $failedRecords += $ideRecord;
      continue;
    }
    if ([string]::IsNullOrEmpty($ideRecord.mlepEmail)){
        Write-Verbose -Message "Missing mlepEmail in CSV record.";
        $ideRecord.Comment = "Missing Email.";
        $recordsNotReadyCount += 1;
        $failedRecords += $ideRecord;
        continue;
    }


          # Build out Group Membership:
      # In the KAMAR files, there is a column containing Group names, separated by '#'
      # Not sure if it is universal, but at least for KAMAR...
      $mlepGroupMembership = $ideRecord.mlepGroupMembership;
      if ($mlepGroupMembership -eq $null){
        if ([string]::IsNullOrWhiteSpace($ideRecord.mlepHomeGroup)){
          $mlepGroupMembership = @();
        }else{
          $mlepGroupMembership = @($ideRecord.mlepHomeGroup);
        }
      }
      else{
        $mlepGroupMembership = ($mlepGroupMembership -as [string]).Split('#', [System.StringSplitOptions]::RemoveEmptyEntries);
      }
      $mlepGroupMembership = $mlepGroupMembership.ForEach({$_.Trim()});     


    Write-Verbose -Message "...Seeing if user already exists...";
    $userExists = $false;
    $Error.Clear();
    $newUser = Get-MsolUser -UserPrincipalName  $ideRecord.mlepEmail -ErrorAction SilentlyContinue 
    if ($Error.Count -eq 0){
      $userExists = $true;
      Write-Verbose -Message "...user already exists. Continuing to next record.";
      # user already exists. This is where we *could* update the existing record, but for now:
      $ideRecord.Comment = "...Already Exists.";
      $recordsAlreadyExistingCount += 1;
      # do not add to either success or error list.
      # $failedRecords += $ideRecord;
      # continue;
    }




    if ($userExists -eq $false){
      Write-Verbose -Message "...user did not exist. Time to create AD entry...";
      Write-Verbose -Message "...Create new AD entry...";
      # sort out first, last and display:   
      Write-Verbose -Message "...cleaning up mlepFirstName...$($ideRecord.mlepFirstName)";
      if (![string]::IsNullOrEmpty($ideRecord.mlepFirstName)){$ideRecord.mlepFirstName = $ideRecord.mlepFirstName.substring(0,1).toupper()+$ideRecord.mlepFirstName.substring(1).tolower();}
      Write-Verbose -Message "...cleaning up mlepLastName...$($ideRecord.mlepLastName)";
      if (![string]::IsNullOrEmpty($ideRecord.mlepLastName)){$ideRecord.mlepLastName = $ideRecord.mlepLastName.substring(0,1).toupper()+$ideRecord.mlepLastName.substring(1).tolower();}
      Write-Verbose -Message "...creating DisplayName...";
      $ideRecord.DisplayName = ($ideRecord.mlepFirstName, $ideRecord.mlepLastName -join $displayNameSeparator).Trim();
      Write-Verbose -Message "...DisplayName is now $($ideRecord.DisplayName)...";

      # Set up password field:   
      # embed password for final report
      Write-Verbose -Message "...cleaning up user's password....";
      if ([string]::IsNullOrWhiteSpace($ideRecord.Password)){$ideRecord.Password = $newADUserDefaultPassword;} 
      if ([string]::IsNullOrWhiteSpace($ideRecord.Password)){
          [string[]]$capitals=$NULL;For ($a=65;$a –le 90;$a++) {$capitals+=,[char][byte]$a }
          [string[]]$lowercase=$NULL;For ($a=97;$a –le 122;$a++) {$lowercase+=,[char][byte]$a }
          [string[]]$numbers=$NULL;For ($a=48;$a –le 57;$a++) {$numbers+=,[char][byte]$a }
          [string[]]$more += [char][byte]64;$more += [char][byte]33;
      
          $ideRecord.Password = ($capitals | GET-RANDOM);
          $ideRecord.Password += ($more | GET-RANDOM)
          For ($i=0; $i –lt 4; $i++) {$ideRecord.Password += ($lowercase | GET-RANDOM);}
          For ($i=0; $i –lt 2; $i++) {$ideRecord.Password += ($numbers | GET-RANDOM);}
      }
      Write-Verbose -Message "...Password for  $($ideRecord.DisplayName) is $($ideRecord.Password)";


      # If Student, and a member of a 'young' group, then the password can be sticky:
      if ($ideRecord.mlepRole -eq "Student"){
        $tmp = ($forceKeepPasswordYears -contains $ideRecord.mlepHomeGroup);
        $ideRecord.ForceKeepPassword =  $tmp;
        # The first entry is the same as group names:
        if ($ideRecord.ForceKeepPassword -eq $false) {
          # If still not a member of Year 1/etc. to which the password must be 
          # fixed, take the time to look through all groups for 
          # Membership:
          foreach ($tmp In $forceKeepPasswordYears){
            if ($mlepGroupMembership -contains $tmp){
              Write-Verbose "...Entry is determined as being a member of '-forceKeepPasswordYears'."
              $ideRecord.ForceKeepPassword = $true; 
            }
          }
        }
      }


      # create new user with the general or specific password.
      Write-Verbose "Creating new user...";
      $Error.Clear();
      $tmp = !$ideRecord.ForceKeepPassword;

      $newUser = New-msolUser `
          -UserPrincipalName $ideRecord.mlepEmail.Trim() `
          -FirstName $ideRecord.mlepFirstName.Trim() `
          -LastName $ideRecord.mlepLastName.Trim() `
          -DisplayName $ideRecord.DisplayName.Trim() `
          -Password $ideRecord.Password `
          -ForceChangePassword $tmp `
          -ErrorAction SilentlyContinue ;

      if ($Error.Count -gt 0){
        Write-Host -Message $Error[0].Exception.Message;
        $ideRecord.Comment = $Error[0]
        $recordsFailedCount += 1;
        $failedRecords += $ideRecord;
      }else{
        $userExists = $true;
        Write-Verbose -Message "...New user created.";
        $ideRecord.Success = $true;
        $recordsCreatedCount+=1;
        $successRecords += $ideRecord;
      }#~else
    }#~if $newUser was false



    # Now that user preexisted or was just created, add to groups:
    if ($userExists -eq $true){

      Write-Verbose -Message "...Adding User to groups...";
      $userId = $newUser.ObjectId;

      #$group = Get-MsolGroup -SearchString:$ideRecord.mlepHomeGroup;

      #if ($ideRecord.mlepHomeGroup -eq $null){
      #  Write-Verbose -Message "...Group column was empty - so won't be adding user to group that is not defined.";
      #  continue;
      #}

      # Write-Verbose -Message "...getting groups that match the 'mlepHomeGroup'..."
      # [PSObject[]] $groups = $groups | Where-Object { ([string]::CompareOrdinal($_.DisplayName, $ideRecord.mlepHomeGroup) -eq 0)}
      # if ($groups.length -eq 0){
      #   Write-Verbose -Message "...none found, so creating new group $($ideRecord.mlepHomeGroup)...";
      #   $group  = New-MsolGroup -DisplayName $ideRecord.mlepHomeGroup -Description "...";
      # }else{
      #   Write-Verbose -Message "...group found...";
      #   $group = $groups[0];
      # }
      # Write-Verbose -Message "...Adding User to Group...";
      # Add-MsolGroupMember -GroupObjectId $group.ObjectId -GroupMemberObjectId $newUser.ObjectId -GroupMemberType:"User" 
      # Write-Verbose -Message "...Done.";

      # On the assumption that HomeGroup is always a member of the option mlepGroupMembership variable:
      Write-Host $mlepGroupMembership;
      [PSObject[]] $groupObjects = Get-MsolGroup;
      foreach ($tmp in $mlepGroupMembership) {
        
        [PSObject[]] $groups = $groupObjects | Where-Object { ([string]::Compare($_.DisplayName, $tmp, $true) -eq 0)}
        if ($groups.length -eq 0){
          Write-Verbose -Message "...none found, so creating new group '$tmp'...";
          $group  = New-MsolGroup -DisplayName $tmp -Description "...";
        }else{
          Write-Verbose -Message "...'$tmp' group found...";
          $group = $groups[0];
        }
        Write-Verbose -Message "...Ensuring User is member of Group...$($group.objectId) : $($newUser.ObjectId)";
        Add-MsolGroupMember -GroupObjectId $group.ObjectId -GroupMemberObjectId $newUser.ObjectId -GroupMemberType:"User"  -ErrorAction SilentlyContinue 
        Write-Verbose -Message "...Done.";
      }#~foreach 
    }#~if userExists

  } #~foreach record



  #$recordsCreatedCount = $successRecords.Count;

  Write-Host "In CSV:${recordsFoundCount}, Processed: ${recordsProcessedCount}, NotReady: ${recordsNotReadyCount} UsersAlreadyExisted: ${recordsAlreadyExistingCount}, Failed: ${recordsFailedCount}, NewlyCreated: ${recordsCreatedCount}";
  Write-Host "Report of what occurred is available in Reports.csv file that was created.";
  
    $tmpArray = @();
    $tmpArray += $successRecords;
    $tmpArray +=  $failedRecords;
    $reportItems = @();
    # sum up the successful ones first:
    foreach($i in $tmpArray){
       $tmp = New-Object PSObject
       Add-Member -InputObject $tmp -MemberType NoteProperty -Name "DateProcessed" -Value $i.DateProcessed;
       Add-Member -InputObject $tmp -MemberType NoteProperty -Name "FirstName" -Value $i.mlepFirstName;
       Add-Member -InputObject $tmp -MemberType NoteProperty -Name "LastName" -Value $i.mlepLastName;
       Add-Member -InputObject $tmp -MemberType NoteProperty -Name "DisplayName" -Value $i.DisplayName;
       Add-Member -InputObject $tmp -MemberType NoteProperty -Name "PrincipalName" -Value $i.mlepEmail;
       Add-Member -InputObject $tmp -MemberType NoteProperty -Name "Password" -Value $i.Password;
       Add-Member -InputObject $tmp -MemberType NoteProperty -Name "ForceKeepPassword" -Value $i.ForceKeepPassword;
       Add-Member -InputObject $tmp -MemberType NoteProperty -Name "Success" -Value $i.Success;
       Add-Member -InputObject $tmp -MemberType NoteProperty -Name "Comment" -Value $i.Comment;
       $reportItems += $tmp;
    }

  $reportItems | Export-CSV  -Path:"./Report.CSV" -Delimiter:$reportColumnSeparator -notype -Append -encoding:"unicode" -Force
         
  return;
}




                          