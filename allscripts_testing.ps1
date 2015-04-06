##################################################################################
#
#
#  Script name: Get-PEFolder.ps1
#
#  Author:      	meadorjc@gmail.com
#
##################################################################################

#param([string]$Name, [switch]$help)

[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint")

function SPHelp([string]$topic) {
    $response = @"
		Add-PEContributeGroup ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]GroupName, [string]Folder)
		Add-PEContributeUser ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]User, [string]Folder)
		Add-PEFolder ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]ParentFolder, [string]NewFolder )
		Add-PEFolderPermission ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]AddTo_Folder, [Microsoft.SharePoint.SPRoleAssignment]SPRoleAssignment)
		Add-PEUserToGroup ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]UserName, [string]GroupName)
		AddCtoGroup ([string]group_name )
		BreakRoleInheritance ([Microsoft.SharePoint.SPListItem]FolderItem, [string]override="D")
		Check-FoldersForObjPermissions ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]ListName, [string]UserOrGroupName)
		Check-ObjHasFolderPermissions ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]EmailOrGroupName, [string]Folder)
		Copy-PEFolderPermissions ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]CopyFrom_Folder, [string]CopyTo_Folder)
		Create-MimePEFolder ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]ParentFolder, [string]CopyPermissions_Folder, [string]NewFolder)
		Delete-PEFolderPermissions ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]DeleteFrom_Folder)
		DisposeWeb (OpenWeb)
		Get-PEFolder ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]Name)
		Get-PEFolderPermissions ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]CopyFrom_Folder)
		Get-PEUserByName ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]UserName)
		Get-SPSite ([string]url)
		Get-SPWeb ([string]url)
		Get-UserOrGroupObj ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]EmailOrGroupName)
		Get-UserSiteGroups ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]Name)
		InitalizeWeb ([string]url)
		Replace-PEFolderPermissions ([Microsoft.SharePoint.SPWeb]OpenWeb, [string]CopyFrom_Folder, [string]CopyTo_Folder)
		ResetRoleInheritance ([Microsoft.SharePoint.SPListItem]FolderItem, [string]override="D")
		TraverseFoldersDump ([Microsoft.SharePoint.SPFolder]FolderObject)
		WriteChildParentFolderPermissions ([Microsoft.SharePoint.SPListItem]self)
"@
	write-host $respose    
    
}


function Get-SPSite([string]$url) {

	return New-Object Microsoft.SharePoint.SPSite($url)
}

function Get-SPWeb([string]$url) {

	$SPSite = Get-SPSite $url
	return $SPSite.OpenWeb()
	$SPSite.Dispose()
}

function Get-PEFolder([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$Name) {

		
	# 6 = Performance Evaluations
	foreach ($Folder in $OpenWeb.Lists["Performance Evaluation"].Folders)
	{
		if ($Folder.name -eq $Name)
		{
			$TargetFolder = $Folder
			break
		}
	
	}

	return $TargetFolder
}

function Add-PEFolder([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$ParentFolder, [string]$NewFolder ){
    
    $ParentFolderObj = Get-PEFolder $OpenWeb $ParentFolder
    
    if ($ParentFolderObj -ne $null) {
        $Folder = $ParentFolderObj.folder.subfolders.Add($NewFolder)
        $Folder.Update()
    }
    return $Folder
}
#return a list of roleassignment objects for specified folder
function Get-PEFolderPermissions([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$CopyFrom_Folder){
    
    $RoleAssigns = @()
    
    $CopyFrom_FolderObj = Get-PEFolder $OpenWeb $CopyFrom_Folder
    
    foreach ($RoleAssignment in $CopyFrom_FolderObj.RoleAssignments){
        $RoleAssigns += $RoleAssignment
    }
    
    return $RoleAssigns
    
}

function Delete-PEFolderPermissions([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$DeleteFrom_Folder){

    $DeleteFrom_FolderObj = Get-PEFolder $OpenWeb $DeleteFrom_Folder
    $DeleteFrom_FolderObj.BreakRoleInheritance($false)
    
    $RoleCount = $DeleteFrom_FolderObj.RoleAssignments.count
    for($i = $RoleCount-1; $i -ge 0; $i -= 1){
        $DeleteFrom_FolderObj.RoleAssignments.Remove($i)
    }
    
    $DeleteFrom_FolderObj.update()
    return $DeleteFrom_FolderObj
}


function Add-PEFolderPermission([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$AddTo_Folder, [Microsoft.SharePoint.SPRoleAssignment]$SPRoleAssignment){
      
    $AddTo_FolderObj = Get-PEFolder $OpenWeb $AddTo_Folder
    
    $AddTo_FolderObj.RoleAssignments.Add($SPRoleAssignment)
    
    $AddTo_FolderObj.update()
    
    return $AddTo_FolderObj
}

function Copy-PEFolderPermissions([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$CopyFrom_Folder, [string]$CopyTo_Folder){
    
    $FolderObj = Get-PEFolder $OpenWeb $CopyTo_Folder
    $UpdatedFolder = $null
    if(BreakRoleInheritance $FolderObj){
        $RoleAssignments = Get-PEFolderPermissions $OpenWeb $CopyFrom_Folder
        
        foreach ($RoleAssignment in $RoleAssignments){
            $UpdatedFolder = Add-PEFolderPermission $OpenWeb $CopyTo_Folder $RoleAssignment
            $UpdatedFolder.Update()
        }
    }
    return $UpdatedFolder
}

function Replace-PEFolderPermissions([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$CopyFrom_Folder, [string]$CopyTo_Folder){
    $FolderObj = Get-PEFolder $OpenWeb $CopyTo_Folder  
    $CopyTo_FolderObj = $null
    if(BreakRoleInheritance $FolderObj){
        $RoleAssignments = Get-PEFolderPermissions $OpenWeb $CopyFrom_Folder
        
        $CopyTo_FolderObj = Delete-PEFolderPermissions $OpenWeb $CopyTo_Folder
        
        foreach ($RoleAssignment in $RoleAssignments){
            $CopyTo_FolderObj = Add-PEFolderPermission $OpenWeb $CopyTo_Folder $RoleAssignment
            $CopyTo_FolderObj.Update()
        }
    }
    return $CopyTo_FolderObj
}   

function Create-MimePEFolder([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$ParentFolder, [string]$CopyPermissions_Folder, [string]$NewFolder){

        $NewFolderObj = Add-PEFolder $OpenWeb $ParentFolder $NewFolder
        
        write-host "Added new folder $($NewFolder)"
        
        if ($NewFolderObj -ne $null){
            $NewFolderObj = Replace-PEFolderPermissions $OpenWeb $CopyPermissions_Folder $NewFolder
            write-host "Replaced $($NewFolder) permissions with the following permissions: $($NewFolderObj.RoleAssignments)"
        }
        
        
        return $NewFolderObj
}


function Get-UserSiteGroups([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$Name) {

 	$Group_Name_List = @()
	
	$SiteGroups = $OpenWeb.SiteGroups
	foreach ($Group in $SiteGroups)
	{
		foreach($User in $Group.Users)
		{
			if (Select-String -inputobject $User.name -pattern $Name)
			{
				$Group_Name_List += new-object PSObject -Property @{ Group = $Group; Name = $User.name}
            }
		}
	}
	

	$OpenWeb.Dispose()
	return $Group_Name_List
}
#THIS IS RECURSIVE
#Get all folders and their members who have a permission to that file and dump in a "|" delimited string
function Dump-FoldersView([Microsoft.SharePoint.SPFolder]$FolderObject){
        if ($FolderObject.SubFolders)
        {
            foreach($sf in $FolderObject.subfolders)
            {
                $dict_list = TraverseFoldersDump $sf
            }
        }
        $member_list = @()
        foreach ($RoleAssignmentObject in $FolderObject.item.RoleAssignments) {
            
            $member_list += $RoleAssignmentObject.member.name
        
        }
        [array]::sort($member_list)
        $reform = ""
        $reform += $($member_list | %{ return $_+"|"})
        return "$($FolderObject.Item.URL)|$($reform)~" + $dict_list
        
}
#THIS IS RECURSIVE
#Get all folders and their members who have a permission to that file and dump in a "|" delimited string
function Dump-MinimalFoldersView([Microsoft.SharePoint.SPFolder]$FolderObject){
        $output = _helper_Dump-MinimalFoldersView $FolderObject
        
        #separate into lines
        $output_split = $output.split("~")
        
        #split and sort outputs (makes it pretty)
        $url_list = @()
        $member_list = @()
        $index = 0                     
        foreach($line in $output_split){
            if ($line -ne $null){
                $line_split = $line.split(":")
                $url_list += "$($line_split[0]).$($index)"
                $member_list += $line_split[1]
            }
            $index++
            
        }
        [array]::sort($url_list)
                        
        #count tabs to put before each name and get max tab count
        $formatted_urls = @()
        $tab_count_list = @()
        $tab_max = 0 #track for later
        foreach($url in $url_list){        
            $url_split = $url.split("/")
            $url_split_count = $url_split.count
            
            $format_tabs = ""
            $tab_count = 0
            for ($i = 0; $i -lt $url_split_count-1; $i++){ $format_tabs += "`t"; $tab_count++}
            if ($tab_count -gt $tab_max){ $tab_max = $tab_count }
            $tab_count_list += $tab_count
            
            #add to formatted list
            $formatted_urls += $format_tabs + $url_split[$url_split_count-1]
        }

        #this is to reformat the line
        $formatted_out = ""
        
        for($i = 0; $i -lt $formatted_urls.count-1; $i++){
            #create tab padding
            $tab_count = 0
            $format_tabs = ""

            for ($j = 0; $j -lt $tab_max-$tab_count_list[$i]; $j++){ $format_tabs += "`t" }
            
            
            $final_split_url = $formatted_urls[$i].split(".")
            $final_url = $final_split_url[0]
            $original_index = $final_split_url[1]
            $formatted_members = $format_tabs + $member_list[$original_index]
            
            $formatted_out += "$($final_url)|$($formatted_members)~"
        
        }
        return $formatted_out
}

function _helper_Dump-MinimalFoldersView([Microsoft.SharePoint.SPFolder]$FolderObject){
        if ($FolderObject.SubFolders)
        {
            foreach($sf in $FolderObject.subfolders)
            {
                $return_list = _helper_Dump-MinimalFoldersView $sf 
            }
        }
        $reform_member_list = ""
        if($FolderObject.name.IndexOf("(") -eq -1){
            $member_list = @()
            foreach ($RoleAssignmentObject in $FolderObject.item.RoleAssignments) {
                
                $member_list += $RoleAssignmentObject.member.name
            
            }
            [array]::sort($member_list)
            
        
            
            $reform_member_list += $($member_list | %{ return $_+"|"})
        }
        return "$($FolderObject.URL):$($reform_member_list)~" + $return_list

}

#Return list of folderitems
function Get-PEFoldersRecursively([Microsoft.SharePoint.SPFolder]$FolderObject){
        
        if ($FolderObject.SubFolders)
        {
            $FolderItemArr = @()    
            foreach($sf in $FolderObject.subfolders)
            {
                $FolderItemArr += Get-PEFoldersRecursively $sf
            }
             
        }
        $FolderItemArr += $FolderObject.Item                
        return $FolderItemArr
        
}
function InitalizeWeb([string]$url){

	return Get-SPWeb $url

}
function DisposeWeb($OpenWeb){

	$OpenWeb.Dispose()
}
#Compare folderitem rolepermissions and parent item role permissions
function WriteChildParentFolderPermissions([Microsoft.SharePoint.SPListItem]$self){
			Write-host "*****CURRENT PERMISSIONS*****" -f BLUE
			$self.RoleAssignments | out-string
            Write-host "*****CURRENT PERMISSIONS*****" -f BLUE   
			
            Write-host
            Write-host "*****PARENT PERMISSIONS (INHERITED)*****" -f BLUE
            $self.folder.ParentFolder.Item.RoleAssignments | out-string    
            Write-host "*****PARENT PERMISSIONS (INHERITED)*****" -f BLUE
            write-host

}


#Override provides a way to bypass prompt with "Y" and "V" for batch jobs, but defaults to prompt
#Can pass each SPistItem from a foreach-loop and Y or V for batch jobs. V only views and doesn't reset. Y resets. If you pass "N" in a batch job
# the function becomes useless 
function ResetRoleInheritance([Microsoft.SharePoint.SPListItem]$FolderItem, [string]$override="D") {

	if($FolderItem.HasUniqueRoleAssignments) {
		
		if ($override -eq "D"){
    		$response = ""
    		Do{
    		    write-host "$($FolderItem.name) has unique permissions status." -f Red	
    			$response = Read-Host "Do you want to reset permissions to inherited status - Yes, No, View? [Y|N|V] " 
    			
    			if ($response -eq "V") {
                    WriteChildParentFolderPermissions $FolderItem
                    read-host "Press any key to continue"
                    
              }
            } until ($response -eq "Y" -or $response -eq "N")
            
        		if ($response -eq "Y"){
        			
        			$FolderItem.ResetRoleInheritance()
        			write-host "$($FolderItem.name)'s permissions have been set to inherit" -f green
        		}
        		elseif ($reponse -eq "N"){
        			
        			write-host "$($FolderItem.name)'s permissions remain unique" -f green
        		}
         }
         elseif ($override -eq "Y"){
         
            $FolderItem.ResetRoleInheritance()
			write-host "$($FolderItem.name)'s permissions have been set to inherit" -f green
         
         }
         elseif ($override -eq "V") {
             	    write-host "$($FolderItem.name) has unique role assignments: $($FolderItem.HasUniqueRoleAssignments)"
                    WriteChildParentFolderPermissions $FolderItem
         }
            
	}
	else {
		write-host "$($FolderItem.name) already has inherited permissions." -f Green
	}
}
function BreakRoleInheritance([Microsoft.SharePoint.SPListItem]$FolderItem, [string]$override="D") {
    
    $Success = $False    
	
    if(!$FolderItem.HasUniqueRoleAssignments) {
		
		if ($override -eq "D"){
    		$response = ""
    		Do{
    		    write-host "$($FolderItem.name) has inherited permissions status." -f Red	
    			$response = Read-Host "Do you want to break inherited status and change to unique status? - Yes, No, View? [Y|N|V] " 
    			
    			if ($response -eq "V") {
                    WriteChildParentFolderPermissions $FolderItem
                    read-host "Press any key to continue"
                    
              }
            } until ($response -eq "Y" -or $response -eq "N")
            
        		if ($response -eq "Y"){
        			
        			$FolderItem.BreakRoleInheritance($true)
        			write-host "$($FolderItem.name)'s permissions have been set to unique" -f green
                    $Success = $True
        		}
        		elseif ($reponse -eq "N"){
        			
        			write-host "$($FolderItem.name)'s permissions remain inherited" -f green
                    $Success = $False
        		}
         }
         elseif ($override -eq "Y"){
         
            $FolderItem.BreakRoleInheritance($false)
			write-host "$($FolderItem.name)'s permissions have been set to unique" -f green
            $Success = $True
         
         }
         elseif ($override -eq "V") {
             	    write-host "$($FolderItem.name) has unique role assignments: $($FolderItem.HasUniqueRoleAssignments)"
                    WriteChildParentFolderPermissions $FolderItem
         }
            
	}
	else {
		write-host "$($FolderItem.name) already has unique permissions." -f Green
        $Success = $True
	}
    return $Success
}

function Get-PEUserByName([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$UserName){
    $User_Name_list = @()
    
    foreach ($user in $OpenWeb.allusers){
		
        if ($user.name -eq $username)
        {
            $TargetUser = $user
        }
        elseif (Select-String -inputobject $User.name -pattern $UserName)
		{
    		$User_Name_list += new-object PSObject -Property @{ Group = $Group; Name = $User.name}
        }
        
    }
    if ($TargetUser -eq $null){
        Write-host "Exact user not found"
        write-host $User_Name_list
    }
    return $TargetUser
}
#Avoid using this...adding individual users as direct permissions is bad! Add Users to site groups instead
function Add-PEContributeUser([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$User, [string]$Folder){
    $FolderObjList = @()
    $FolderObjList += Get-PEFolder $OpenWeb $Folder
    $ParentFolderList = $FolderObjList[0].URL.split("/")
    #URL format == "Performance Evaluation/ParentFolder/Folder"
    #i begins 1 because arr[0] == "Performance Evaluation" list, not folder
    #i ends at count-1 because we already have the main folder object
    for($i = 1; $i -lt $ParentFolderList.count-1; $i++){
        $FolderObjList += Get-PEFolder $OpenWeb $ParentFolderList[$i]
    }
    
    $UserObj = Get-PEUserByName $OpenWeb $User
    $Success = $False
    foreach($FolderObj in $FolderObjList){        
        if($FolderObj){
            if(BreakRoleInheritance $FolderObj){
            
                if($UserObj -ne $null){
                    $NewRole = New-Object Microsoft.SharePoint.SPRoleAssignment $UserObj 
                    $NewRole.RoleDefinitionBindings.Add($OpenWeb.RoleDefinitions["Contribute"])
                }
                else{
                    write-host "User name ($($User)) could not be found"
                }
                if($FolderObj -ne $null){
                    $FolderObj.RoleAssignments.Add($NewRole)
                    $Success = $True
                }
                else{
                    write-host "Folder name ($($Folder)) could not be found"
                }
           }
        }
    }
    return $FolderObjList
}



#Add test account to a specific group and remove from all other groups
function AddCtoGroup([string]$group_name ) {
    $ctest = $r.AllUsers | %{if ($_.name -eq "tstcaleb tstmeador"){ return $_ }}
    if ($ctest.name -ne "tstcaleb tstmeador") {
        write-host "User is not test account"
    
    }
    $add_group = $r.sitegroups[$group_name]
    foreach ($group in $ctest.groups)
    {
        $group.users.remove($ctest)
    }
    if ($add_group)
    {
        $add_group.users.add($ctest.loginname, $ctest.name, $ctest.email, $ctest.notes)
    }
    else { write-host "Group does not exist"
    }
}
function Add-PEUserToGroup([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$UserName, [string]$GroupName){
    $UserNameObj = Get-PEUserByName $OpenWeb $UserName
    $AddGroupObj = $OpenWeb.sitegroups[$GroupName]
    $Success = $False
    
    if ($AddGroupObj -and $UserNameObj)
    {
        $AddGroupObj.users.add($UserNameObj.loginname, $UserNameObj.name, $UserNameObj.email, $UserNameObj.notes)
        $Success = $True
    }
    if($AddGroupObj -eq $null){ 
        write-host "Cannot locate group"
        $Success = $False
    }
    if($UserNameObj -eq $null){
        write-host "Cannot locate user"
        $Success = $False
    }
    return $Success
}
#This will iterate up the Folder structure to parent folders to ensure vision of subfolders by group!
function Add-PEContributeGroup([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$GroupName, [string]$Folder){
    $FolderObjList = @()
    $SuccessList = @{}
    $FolderObjList += Get-PEFolder $OpenWeb $Folder
    $ParentFolderList = $FolderObjList[0].URL.split("/")
    #URL format == "Performance Evaluation/ParentFolder/Folder"
    #i begins 1 because arr[0] == "Performance Evaluation" list, not folder
    #i ends at count-1 because we already have the main folder object
    for($i = 1; $i -lt $ParentFolderList.count-1; $i++){
        $FolderObjList += Get-PEFolder $OpenWeb $ParentFolderList[$i]
    }
    #$FolderObjList | %{write-host $_.name}
    
    $GroupObj = $OpenWeb.sitegroups[$GroupName]
    
    foreach($FolderObj in $FolderObjList){
        $PermissionTable = @{}
        $FolderObj.RoleAssignments | %{ $PermissionTable += @{$_.Member.Name=$_.RoleDefinitionBindings | %{$_.name}}}
        
        if($GroupObj -ne $null){
					if(!$($PermissionTable[$GroupObj.Name] -contains "Contribute")){
						if(BreakRoleInheritance $FolderObj){
							$NewRole = New-Object Microsoft.SharePoint.SPRoleAssignment $GroupObj 
							$NewRole.RoleDefinitionBindings.Add($OpenWeb.RoleDefinitions["Contribute"])
			
							if ($FolderObj -ne $null){ 
									$FolderObj.RoleAssignments.Add($NewRole)
									write-host "$($GroupObj.Name) = Contribute on $($FolderObj.Name)"
							}
							else { 
								write-host "Folder name ($($Folder)) could not be found"
								"$($FolderObj.Name) not found" 
							}
						}
					}
					else{ write-host "$($GroupObj.Name) already has contribute permission to $($FolderObj.Name)" } 
				}		
				else{ write-host "Group name ($($GroupName)) could not be found" }
		}
    return $SuccessList
}

#Have to use by user email@domain.org or groupname
function Get-UserOrGroupObj([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$EmailOrGroupName){
    try{
        $UserOrGroupObj = $OpenWeb.AllUsers.GetByEmail($EmailOrGroupName)
    }
    catch{
        $UserOrGroupObj = $Null
    }
    if($UserOrGroupObj -eq $Null){
        try{
            $UserOrGroupObj = $OpenWeb.sitegroups[$EmailOrGroupName]
        }
        catch{
            $UserOrGroupObj = $Null
        }
    }
        
    return $UserOrGroupObj
}

#Return RoleAssignment if name found else return $null; Must pass email@domain.org or group name
function Check-ObjHasFolderPermissions([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$EmailOrGroupName, [string]$Folder){
    $UserOrGroupObj = Get-UserOrGroupObj $OpenWeb $UserOrGroupName
    
    if ($UserOrGroupObj -ne $Null) {
        $FolderObj = Get-PEFolder $OpenWeb $Folder
    }
       
    $RoleAssignment = $Null
    foreach ($RoleAssignment in $FolderObj.RoleAssignments) {
        if ($RoleAssignment.Member.Name -eq $UserOrGroupObj.Name){
                $FoundRole = $RoleAssignment
                break    
        }
    }
    return $RoleAssignment
}
#Return RoleAssignment if name found else return $null; Must pass email@domain.org or group name
function Check-GroupHasFolderPermissions([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$GroupName, [string]$Folder){
    $FolderObj = Get-PEFolder $OpenWeb $Folder
    
    $RoleAssignment = $Null
    $FoundRole = $null
    $Roles = $null
    foreach ($RoleAssignment in $FolderObj.RoleAssignments) {
        if ($RoleAssignment.Member.Name -eq $GroupName){
                $roles = "{"
                foreach($RoleDefinitionBinding in $RoleAssignment.RoleDefinitionBindings){
                    $roles += " $($RoleDefinitionBinding.Name)"
                }
                $roles += " }"
                $FoundRole = $RoleAssignment.Member.Name
                break    
        }
    }
    if ($FoundRole -eq $null){
        $Roles = $null
    }
    return $Roles
}
#List name should be the name of the document library; ie "Performance Evaluation" in COI case; 
function Check-FoldersForObjPermissions([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$ListName, [string]$UserOrGroupName){
    $ListObj = $OpenWeb.Lists[$ListName]
        
    $PermissionList = @()
    foreach ($FolderItemObj in $ListObj.Folders) {
        $Permission = Check-ObjHasFolderPermissions $OpenWeb $UserOrGroupName $FolderItemObj.Name
        if ($Permission -ne $Null){
            $PermissionList += $Permission
        }
    }

    return $PermissionList
}

function QuickFolderView([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$FolderName) {
    $($(TraverseFoldersDump  $(get-pefolder $OpenWeb $FolderName).folder).split("~")) | %{$_.split("|")} | select-string "Perf"
}

if($help) { GetHelp; Continue }
if($Name) { Get-PEFolder -Name $Name }


function Audit-MemberPermissions([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$ListName, [string]$MemberName){
	$AuditTable = @()
    
    #iterate through all folders
    foreach ($Folder in $OpenWeb.Lists[$ListName].folders){
		$t = $false;
		$url = $Folder.url;
        $RoleList = @()
		
        #iterate through folder's role assignments
        foreach($RoleAssignment in $Folder.roleassignments){ 
			if ($RoleAssignment.member.name -eq $MemberName){
				$t = $true;
                $ContributeTest = $False
                #collect roles if true
                foreach($RoleDefinitionBinding in $RoleAssignment.RoleDefinitionBindings){
                    if ($RoleDefinitionBinding.Name -eq "Contribute"){
                        $ContributeTest = $True
                        break
                    }
                }
                
                
			}
		}
     #We're going to add the folders to the list weather permission or not...just change the format
		if($t -eq $true -and $ContributeTest -eq $True){
            
			$AuditTable += "$($Folder.name)|$($Folder.url)|Contribute"
        
			#write-host $url "true; permissions: " -f DarkGreen
		}
        elseif($t -eq $true -and $ContributeTest -eq $False){
            $AuditTable += "$($Folder.name)|$($Folder.url)|None"
            #$obj = @{ Folder = $Folder.name; URL = $Folder.url; Permissions = "NONE"}
            #$AuditTable += New-Object PSObject -Property $obj 
        }
		else{
			$AuditTable += "$($Folder.name)|$($Folder.url)|None"
            #$obj = @{ Folder = $Folder.name; URL = $Folder.url; Permissions = "NONE" }
            #$AuditTable += New-Object PSObject -Property $obj
			#write-host $_.name "false; Adding permissions:" -f Red
			#write-host $_.url $_.name
			
		}
	}
	return $AuditTable
    
}
#Add(string name, Microsoft.SharePoint.SPMember owner, Microsoft.SharePoint.SPUser defaultUser, string description)
function Add-PEPermissionGroup([Microsoft.SharePoint.SPWeb]$OpenWeb, [string]$ListName, [string]$GroupName, [string]$Description){
    $GroupObj = $OpenWeb.sitegroups[$GroupName]
    $Success = $False   
    
        if($GroupObj -eq $null){
            $OpenWeb.sitegroups.Add($GroupName, $OpenWeb.CurrentUser, $OpenWeb.CurrentUser, $Description)
            
            $GroupObj = $OpenWeb.sitegroups[$GroupName]
            if($GroupObj -ne $null){
                $NewRole = New-Object Microsoft.SharePoint.SPRoleAssignment $GroupObj 
                $NewRole.RoleDefinitionBindings.Add($OpenWeb.RoleDefinitions["Contribute"])
                $Openweb.Lists[$ListName].RoleAssignments.Add($NewRole)
                $Succes = $True
            }
            else{
                write-host "Group name ($($GroupName)) could not be created!"
                $Success = $False
            }
        }
        else{
            write-host "Group name ($($GroupName)) already exists!"
            $Success = $False
        }    
        
    
    return $Success
    
}
