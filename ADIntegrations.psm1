function Get-RolePositionGroup {
    [CmdletBinding(SupportsShouldProcess=$false,DefaultParameterSetName="OU")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="The title of the position which we are returning the matchins role group")]
		[Object]$PositionTitle,
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="Active Directory OU to limit search of role group to")]
        [String]$OU,
        [Parameter(Mandatory=$false,HelpMessage="The property to match PositionTitle against")]
        [String]$PropertyName="displayName"
	)
	Process {
        if ($OU) {
            return Get-ADGroup -Filter "$PropertyName -EQ '$PositionTitle'" -SearchBase $OU
        } else {
            return Get-ADGroup -Filter "$PropertyName -EQ '$PositionTitle'"
        }
	}
}

function Get-GroupDoorPermission {
	[CmdletBinding(SupportsShouldProcess=$false,DefaultParameterSetName="OU")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="Identity of Group, same as Get-ADPrincipalGroupMembership")]
		[Object]$Identity,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipeline=$false,ParameterSetName="OU",HelpMessage="Active Directory OU to limit search of permission group to")]
        [String]$OU,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipeline=$false,ParameterSetName="Prefix",HelpMessage="Prefix of AD group name to remove when looking up Access codes")]
        [String]$ADGroupPrefix
	)
	Process {
        if ($OU) {
            return Get-GroupMembershipRecursive -Identity $Identity | Where-Object -Property distinguishedName -Like -Value "*$OU"
        } else {
            return Get-GroupMembershipRecursive -Identity $Identity | Where-Object -Property name -Like -Value "$ADGroupPrefix*"
        }
	}
}


function Get-GroupDoorAccessCode {
	[CmdletBinding(SupportsShouldProcess=$false,DefaultParameterSetName="OU")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="Identity of Group, same as Get-ADPrincipalGroupMembership")]
		[Object]$Identity,
        [Parameter(Mandatory=$false,ParameterSetName="OU",HelpMessage="Active Directory OU to limit search of permission group to")]
        [String]$OU,
        [Parameter(Mandatory=$false,HelpMessage="Prefix of AD group name to remove when looking up Access codes")]
        [String]$ADGroupPrefix,
        [Parameter(Mandatory=$false,HelpMessage="Prefix of AccessCode names to add to AD group name when looking up Access codes")]
        [String]$SMSAccessCodePrefix = "",
        [Parameter(Mandatory=$false,HelpMessage="Limit results to specific CompanyID")]
        [int]$CompanyID,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {
        if ($OU) {
            $Groups = Get-GroupDoorPermission -Identity $Identity -OU $OU
        } else {
            $Groups = Get-GroupDoorPermission -Identity $Identity -ADGroupPrefix $ADGroupPrefix
        }
        $AccessCodes = @()
        forEach ($Group in $Groups) {
            if ($ADGroupPrefix) {
                $AccessCodeName = $SMSAccessCodePrefix + $Group.name.replace($ADGroupPrefix, "")
            } else {
                $AccessCodeName = $SMSAccessCodePrefix + $Group.name
            }
            if ($CompanyID) {
                $AccessCodes += Get-SMSAccessCode -AccessCodeName $AccessCodeName -CompanyID $CompanyID -SMSConnection $SMSConnection
            } else {
                $AccessCodes += Get-SMSAccessCode -AccessCodeName $AccessCodeName -SMSConnection $SMSConnection
            }
        }
        return $AccessCodes
	}
}

Export-ModuleMember -Function *