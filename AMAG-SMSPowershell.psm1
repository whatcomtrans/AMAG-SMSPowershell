function add-BoundParams {
    param(
        [hashtable] $NamedValues,
        [Object] $ToAdd,
        [Array] $Ignore
    )

    forEach ($ParamKey in $ToAdd.Keys) {
        if ($Ignore -notcontains $ParamKey) {
            $NamedValues.Add($ParamKey, $ToAdd.$ParamKey)
        }
    }
}

function New-SMSCommand {
    [CmdletBinding(SupportsShouldProcess=$false)]
	Param(
		[Parameter(Mandatory=$false,Position=0,HelpMessage="Hashtable of Name Value pairs to insert into the DataImportTable")]
		[hashtable] $NamedValues,
        [Parameter(Mandatory=$false,Position=1,HelpMessage="SMSServerConnection object to use.  If not set here, object must be set on the object directly.")]
		[Object] $SMSServerConnection
    )
    
    $scriptSQL = {
        [hashtable] $NamedValues = $this.NamedValues
        [String] $Collumns = ""
        [String] $Values = ""
        [String[]] $SupportedCollumns = @("RecordCount","LastName","FirstName","CardNumber","CompanyID","CardIssueLevel","EmployeeReference","PIN","PersonalData1","PersonalData2","PersonalData3","PersonalData4","PersonalData5","PersonalData6","PersonalData7","PersonalData8","PersonalData9","PersonalData10","ActiveDate","ExpiryDate","ReaderGroupID","TimeCodeID","RecordRequest","RecordStatus","InactiveComment","Encryption","CustomerCode","FaceFile","SignatureFile","InitLet","BadgeFormatID","PersonalData11","PersonalData12","PersonalData13","PersonalData14","PersonalData15","PersonalData16","PersonalData17","PersonalData18","PersonalData19","PersonalData20","PersonalData21","PersonalData22","PersonalData23","PersonalData24","PersonalData25","PersonalData26","PersonalData27","PersonalData28","PersonalData29","PersonalData30","PersonalData31","PersonalData32","PersonalData33","PersonalData34","PersonalData35","PersonalData36","PersonalData37","PersonalData38","PersonalData39","PersonalData40","PersonalData41","PersonalData42","PersonalData43","PersonalData44","PersonalData45","PersonalData46","PersonalData47","PersonalData48","PersonalData49","PersonalData50","HandTemplateValue1","HandTemplateValue2","HandTemplateValue3","HandTemplateValue4","HandTemplateValue5","HandTemplateValue6","HandTemplateValue7","HandTemplateValue8","HandTemplateValue9","ReaderID","AccessCodeID","ImportNow","BatchReference","DefaultBadge","IDSCode","AreaID","DeactivateAtThreatLevel","CardUsageRemaining","Priority")

        #Iterate through values
        if ($NamedValues) {
            forEach ($key in $NamedValues.Keys) {
                #Check NamedValues that are not compatible with DataImportTable
                if ($SupportedCollumns -contains $key) {
                    $Collumns += ",$key"
                    switch (($NamedValues.$key).GetTypeCode()) {
                        "String" {$Values += (" ,'" + $NamedValues.$key + "'")}
                        "DateTime" {$Values += (" ,'" + ([DateTime] ($NamedValues.$key)).ToString() + "'")}
                        default {$Values += (" ,'" + $NamedValues.$key + "'")}
                    }
                }
            }
        }
        
        $Collumns = $Collumns.Remove(0,1)
        $Values = $Values.Remove(0,2)
    
        [String] $SQL = "DECLARE @MyTableVar table(RecordCount int); INSERT INTO DataImportTable ($Collumns) OUTPUT INSERTED.RecordCount INTO @MyTableVar VALUES ($Values); SELECT RecordCount FROM @MyTableVar"
        return $SQL
    }

    $scriptExecute = {
        $SMSConnection = $this.SMSServerConnection
        $SQLCommand = $this.SQLCommand
        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        $this.RecordCount = $retvalue.RecordCount
    }

    $scriptRecordStatus = {
        $SMSConnection = $this.SMSServerConnection
        $SQLCommand = "SELECT RecordStatus FROM DataImportTable WHERE RecordCount = " + $this.RecordCount
        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue.RecordStatus
    }

    $scriptRecordStatusDescription = {
        $SMSConnection = $this.SMSServerConnection
        $SQLCommand = "SELECT Message FROM MessageTable WHERE RecordStatus = " + $this.RecordStatus
        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue.Message
    }

    $scriptAddNamedValue = {
        param(
            [String] $Name,
            [Object] $Value
        )
        
        [hashtable] $nv = $this.NamedValues
        $nv.Add($Name,$Value)
    }

    $SQLObj = New-Object -TypeName PSObject
    $SQLObj | Add-Member ScriptProperty SQLCommand $scriptSQL
    if (!$SMSServerConnection) {
        $SQLObj | Add-Member Noteproperty SMSServerConnection $null
    } else {
        $SQLObj | Add-Member Noteproperty SMSServerConnection $SMSServerConnection
    }
    $SQLObj | Add-Member ScriptMethod Execute $scriptExecute
    $SQLObj | Add-Member ScriptMethod AddNamedValue $scriptAddNamedValue
    $SQLObj | Add-Member ScriptProperty RecordStatus $scriptRecordStatus
    $SQLObj | Add-Member ScriptProperty RecordStatusDescription $scriptRecordStatusDescription
    $SQLObj | Add-Member Noteproperty RecordCount 0
    if (!$NamedValues) {
        $SQLObj | Add-Member Noteproperty NamedValues ([hashtable] @{})
    } else {
        $SQLObj | Add-Member Noteproperty NamedValues $NamedValues
    }
    
    return $SQLObj
}

function Get-SMSServerConnection {
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
		[Parameter(Mandatory=$true,Position=0,HelpMessage="put help here")]
		[String]$SMSDatabaseServer,
        [Parameter(Mandatory=$true,Position=1,HelpMessage="put help here")]
		[String]$SMSImportDatabase,
        [Parameter(Mandatory=$false,Position=2,HelpMessage="put help here")]
		[String]$SMSImportDatabaseUsername,
        [Parameter(Mandatory=$false,Position=3,HelpMessage="put help here")]
		[String]$SMSImportDatabasePassword,
        [Parameter(Mandatory=$false,Position=4,HelpMessage="put help here")]
		[int]$SMSImportIntervalMinutes = 1
	)
	Process {
        $conn = New-Object –TypeName PSObject –Prop (@{
                'SMSDatabaseServer'=$SMSDatabaseServer;
                'SMSImportDatabase'=$SMSImportDatabase;
                'SMSImportDatabaseUsername'=$SMSImportDatabaseUsername;
                'SMSImportDatabasePassword'=$SMSImportDatabasePassword;
                'SMSImportIntervalMinutes'=$SMSImportIntervalMinutes});
        $Global:DefaultSMSServerConnection = $conn
        return $conn
	}
}

function Disable-SMSCard {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="CardNumber")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="CardNumber",HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeNumber",HelpMessage="Employee reference number, typically the employee number or employee ID.  If passed ADUser object, uses EmployeeID")]
        [alias("EmployeeID")]
		[String]$EmployeeReference,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$true,HelpMessage="Customer/Facility Code")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="Use this switch to have the process wait until records are processed (about 1 minute) and return any errors.")]
        [switch]$Wait,
        [Parameter(Mandatory=$false,HelpMessage="Switch to return the SMSCommand object generated and executed by this cmdlet, can be combined with whatif.  Otherwise nothing is returned")]
        [switch]$ReturnSMSCommand,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Begin {
        [Object[]]$SMSCommands = @()
	}
	Process {
        [hashtable] $RecordNamedValues = @{}
        
        $RecordNamedValues.Add("RecordRequest",2)
        $RecordNamedValues.Add("RecordStatus",0)
        $RecordNamedValues.Add("ImportNow",1)
        $RecordNamedValues.Add("CustomerCode",$CustomerCode)
        
        switch ($PsCmdlet.ParameterSetName) {
            "CardNumber" {
                $RecordNamedValues.Add("CardNumber",$CardNumber)
                $Item = $CardNumber
            }
            "EmployeeNumber" {
                $RecordNamedValues.Add("EmployeeReference",$EmployeeReference)
                $Item = $EmployeeReference
            }
            "EmployeeName" {
                $RecordNamedValues.Add("LastName",$LastName)
                $RecordNamedValues.Add("FirstName",$FirstName)
                $Item = "$LastName, $FirstName"
            }
        }
        
        $SMSCommand = (New-SMSCommand -NamedValues $RecordNamedValues -SMSServerConnection $SMSConnection)
        $SMSCommands = $SMSCommands + $SMSCommand
        If ($PSCmdlet.ShouldProcess("$Item","Disable Card")) {
            $SMSCommand.Execute()
        }
        
        if ($ReturnSMSCommand) {
            return $SMSCommand
        }
	}
	End {
        if ($Wait) {
            Sleep (($SMSConnection.SMSImportIntervalMinutes * 60) + 1)
        }
	}
}

function Enable-SMSCard {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="CardNumber")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="CardNumber",HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeNumber",HelpMessage="Employee reference number, typically the employee number or employee ID.  If passed ADUser object, uses EmployeeID")]
        [alias("EmployeeID")]
		[String]$EmployeeReference,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$true,HelpMessage="Customer/Facility Code")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="Use this switch to have the process wait until records are processed (about 1 minute) and return any errors.")]
        [switch]$Wait,
        [Parameter(Mandatory=$false,HelpMessage="Switch to return the SMSCommand object generated and executed by this cmdlet, can be combined with whatif.  Otherwise nothing is returned")]
        [switch]$ReturnSMSCommand,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Begin {
        [Object[]]$SMSCommands = @()
	}
	Process {
        [hashtable] $RecordNamedValues = @{}
        
        $RecordNamedValues.Add("RecordRequest",5)
        $RecordNamedValues.Add("RecordStatus",0)
        $RecordNamedValues.Add("ImportNow",1)
        $RecordNamedValues.Add("CustomerCode",$CustomerCode)
        
        switch ($PsCmdlet.ParameterSetName) {
            "CardNumber" {
                $RecordNamedValues.Add("CardNumber",$CardNumber)
                $Item = $CardNumber
            }
            "EmployeeNumber" {
                $RecordNamedValues.Add("EmployeeReference",$EmployeeReference)
                $Item = $EmployeeReference
            }
            "EmployeeName" {
                $RecordNamedValues.Add("LastName",$LastName)
                $RecordNamedValues.Add("FirstName",$FirstName)
                $Item = "$LastName, $FirstName"
            }
        }
        
        $SMSCommand = (New-SMSCommand -NamedValues $RecordNamedValues -SMSServerConnection $SMSConnection)
        $SMSCommands = $SMSCommands + $SMSCommand
        If ($PSCmdlet.ShouldProcess("$Item","Disable Card")) {
            $SMSCommand.Execute()
        }
        
        if ($ReturnSMSCommand) {
            return $SMSCommand
        }
	}
	End {
        if ($Wait) {
            Sleep (($SMSConnection.SMSImportIntervalMinutes * 60) + 1)
        }
	}
}

function Set-SMSCard {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$false,Position=0,ValueFromPipelineByPropertyName=$true,HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipelineByPropertyName=$true,HelpMessage="Employee reference number, typically the employee number or employee ID.  If passed ADUser object, uses EmployeeID")]
        [alias("EmployeeID")]
		[String]$EmployeeReference,
        [Parameter(Mandatory=$false,Position=2,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$false,Position=3,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders Middle name")]
        [alias("Initials")]
		[String]$MiddleName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Date card is set to become active")]
        [DateTime]$ActiveDate,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Date card is set to become in-active")]
        [alias("accountExpires")]
        [DateTime]$InactiveDate,
        [Parameter(Mandatory=$true,HelpMessage="Customer/Facility Code")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="Other attributes to set on the card as a hashtable of name value pairs using the database field names and data types of the DataImportTable.")]
        [hashtable]$OtherNamedValues,
        [Parameter(Mandatory=$false,HelpMessage="Use this switch to have the process wait until records are processed (about 1 minute) and return any errors.")]
        [switch]$Wait,
        [Parameter(Mandatory=$false,HelpMessage="Switch to return the SMSCommand object generated and executed by this cmdlet, can be combined with whatif.  Otherwise nothing is returned")]
        [switch]$ReturnSMSCommand,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Begin {
        [Object[]]$SMSCommands = @()
	}
	Process {
        [hashtable] $RecordNamedValues = @{}
        
        $RecordNamedValues.Add("RecordRequest",1)
        $RecordNamedValues.Add("RecordStatus",0)
        $RecordNamedValues.Add("ImportNow",1)        
        
        add-BoundParams $RecordNamedValues $PSBoundParameters @("Wait","ReturnSMSCommand","SMSConnection", "OtherNamedValues")
        if ($OtherNamedValues) {
            $RecordNamedValues = $RecordNamedValues + $OtherNamedValues
        }

        $Item = "Making changes to card."
        $SMSCommand = (New-SMSCommand -NamedValues $RecordNamedValues -SMSServerConnection $SMSConnection)
        $SMSCommands = $SMSCommands + $SMSCommand
        If ($PSCmdlet.ShouldProcess("$Item","Modify Card")) {
            $SMSCommand.Execute()
        }
        
        if ($ReturnSMSCommand) {
            return $SMSCommand
        }
	}
	End {
        if ($Wait) {
            Sleep (($SMSConnection.SMSImportIntervalMinutes * 60) + 1)
        }
	}
}

function Add-SMSCard {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipelineByPropertyName=$true,HelpMessage="Employee reference number, typically the employee number or employee ID.  If passed ADUser object, uses EmployeeID")]
        [alias("EmployeeID")]
		[String]$EmployeeReference,
        [Parameter(Mandatory=$true,Position=2,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$true,Position=3,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders Middle name")]
        [alias("Initials", "MiddleName")]
		[String]$InitLet,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Date card is set to become active, defaults to current date.")]
        [DateTime]$ActiveDate,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Date card is set to become in-active")]
        [alias("accountExpires", "InactiveDate")]
        [DateTime]$ExpiryDate,
        [Parameter(Mandatory=$true,HelpMessage="Customer/Facility Code")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="Other attributes to set on the card as a hashtable of name value pairs using the database field names and data types of the DataImportTable.")]
        [hashtable]$OtherNamedValues,
        [Parameter(Mandatory=$false,HelpMessage="Use this switch to have the process wait until records are processed (about 1 minute) and return any errors.")]
        [switch]$Wait,
        [Parameter(Mandatory=$false,HelpMessage="Switch to return the SMSCommand object generated and executed by this cmdlet, can be combined with whatif.  Otherwise nothing is returned")]
        [switch]$ReturnSMSCommand,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Begin {
        [Object[]]$SMSCommands = @()
	}
	Process {
        [hashtable] $RecordNamedValues = @{}
        
        $RecordNamedValues.Add("RecordRequest",0)
        $RecordNamedValues.Add("RecordStatus",0)
        $RecordNamedValues.Add("ImportNow",1)        
        
        add-BoundParams $RecordNamedValues $PSBoundParameters @("Wait","ReturnSMSCommand","SMSConnection", "OtherNamedValues")
        if ($OtherNamedValues) {
            $RecordNamedValues = $RecordNamedValues + $OtherNamedValues
        }

        if (Get-SMSCard -CardNumber $CardNumber -CustomerCodeNumber $CustomerCode -SMSConnection $SMSConnection) {
            throw "Card already exists."
        }

        $Item = "Adding new card number $CardNumber for $FirstName $LastName."
        $SMSCommand = (New-SMSCommand -NamedValues $RecordNamedValues -SMSServerConnection $SMSConnection)
        $SMSCommands = $SMSCommands + $SMSCommand
        If ($PSCmdlet.ShouldProcess("$Item","Modify Card")) {
            $SMSCommand.Execute()
        } else {
            Echo $SMSCommand
        }
        
        if ($ReturnSMSCommand) {
            return $SMSCommand
        }
	}
	End {
        if ($Wait) {
            Sleep (($SMSConnection.SMSImportIntervalMinutes * 60) + 1)
        }
	}
}

function Remove-SMSCard {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="CardNumber")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="CardNumber",HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeNumber",HelpMessage="Employee reference number, typically the employee number or employee ID.  If passed ADUser object, uses EmployeeID")]
        [alias("EmployeeID")]
		[String]$EmployeeReference,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$true,Position=2,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$true,HelpMessage="Customer/Facility Code")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="Use this switch to have the process wait until records are processed (about 1 minute) and return any errors.")]
        [switch]$Wait,
        [Parameter(Mandatory=$false,HelpMessage="Switch to return the SMSCommand object generated and executed by this cmdlet, can be combined with whatif.  Otherwise nothing is returned")]
        [switch]$ReturnSMSCommand,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Begin {
        [Object[]]$SMSCommands = @()
	}
	Process {
        [hashtable] $RecordNamedValues = @{}
        
        $RecordNamedValues.Add("RecordRequest",42)
        $RecordNamedValues.Add("RecordStatus",0)
        $RecordNamedValues.Add("ImportNow",1)        
        
        add-BoundParams $RecordNamedValues $PSBoundParameters @("Wait","ReturnSMSCommand","SMSConnection")
        if ($OtherNamedValues) {
            $RecordNamedValues = $RecordNamedValues + $OtherNamedValues
        }

        $Item = "Adding AccessCode $AccessCodeID."
        $SMSCommand = (New-SMSCommand -NamedValues $RecordNamedValues -SMSServerConnection $SMSConnection)
        $SMSCommands = $SMSCommands + $SMSCommand
        If ($PSCmdlet.ShouldProcess("$Item","Add Access Code")) {
            $SMSCommand.Execute()
        }
        
        if ($ReturnSMSCommand) {
            return $SMSCommand
        }
	}
	End {
        if ($Wait) {
            Sleep (($SMSConnection.SMSImportIntervalMinutes * 60) + 1)
        }
	}
}

function Add-SMSAccessRights {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="CardNumber")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="CardNumber",HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$true,Position=1,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeNumber",HelpMessage="Employee reference number, typically the employee number or employee ID.  If passed ADUser object, uses EmployeeID")]
        [alias("EmployeeID")]
		[String]$EmployeeReference,
        [Parameter(Mandatory=$true,Position=2,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$true,Position=3,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="AccessCodeID to add to card(s) found.")]
        [int]$AccessCodeID,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Customer/Facility Code")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="Use this switch to have the process wait until records are processed (about 1 minute) and return any errors.")]
        [switch]$Wait,
        [Parameter(Mandatory=$false,HelpMessage="Switch to return the SMSCommand object generated and executed by this cmdlet, can be combined with whatif.  Otherwise nothing is returned")]
        [switch]$ReturnSMSCommand,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Begin {
        [Object[]]$SMSCommands = @()
	}
	Process {
        [hashtable] $RecordNamedValues = @{}
        
        $RecordNamedValues.Add("RecordRequest",3)
        $RecordNamedValues.Add("RecordStatus",0)
        $RecordNamedValues.Add("ImportNow",1)        
        
        add-BoundParams $RecordNamedValues $PSBoundParameters @("Wait","ReturnSMSCommand","SMSConnection")
        if ($OtherNamedValues) {
            $RecordNamedValues = $RecordNamedValues + $OtherNamedValues
        }

        $Item = "$CardNumber$EmployeeReference$FirstName $LastName"
        $SMSCommand = (New-SMSCommand -NamedValues $RecordNamedValues -SMSServerConnection $SMSConnection)
        $SMSCommands = $SMSCommands + $SMSCommand
        If ($PSCmdlet.ShouldProcess("$Item","Add Access Code $AccessCodeID")) {
            $SMSCommand.Execute()
        }
        
        if ($ReturnSMSCommand) {
            return $SMSCommand
        }
	}
	End {
        if ($Wait) {
            Sleep (($SMSConnection.SMSImportIntervalMinutes * 60) + 1)
        }
	}
}

function Remove-SMSAccessRights {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="CardNumber")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="CardNumber",HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeNumber",HelpMessage="Employee reference number, typically the employee number or employee ID.  If passed ADUser object, uses EmployeeID")]
        [alias("EmployeeID")]
		[String]$EmployeeReference,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$true,Position=2,ValueFromPipelineByPropertyName=$true,ParameterSetName="EmployeeName",HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$false,HelpMessage="AccessCodeID to remove from card(s) found.  If no AccessCodeID specified, all rights are removed.")]
        [int]$AccessCodeID,
        [Parameter(Mandatory=$true,HelpMessage="Customer/Facility Code")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="Use this switch to have the process wait until records are processed (about 1 minute) and return any errors.")]
        [switch]$Wait,
        [Parameter(Mandatory=$false,HelpMessage="Switch to return the SMSCommand object generated and executed by this cmdlet, can be combined with whatif.  Otherwise nothing is returned")]
        [switch]$ReturnSMSCommand,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Begin {
        [Object[]]$SMSCommands = @()
	}
	Process {
        [hashtable] $RecordNamedValues = @{}
        
        if (!$AccessCodeID) {
            $RecordNamedValues.Add("RecordRequest",4)
        } else {
            $RecordNamedValues.Add("RecordRequest",6)
        }
        $RecordNamedValues.Add("RecordStatus",0)
        $RecordNamedValues.Add("ImportNow",1)        
        
        add-BoundParams $RecordNamedValues $PSBoundParameters @("Wait","ReturnSMSCommand","SMSConnection")
        if ($OtherNamedValues) {
            $RecordNamedValues = $RecordNamedValues + $OtherNamedValues
        }

        $Item = "$CardNumber$EmployeeReference$FirstName $LastName"
        $SMSCommand = (New-SMSCommand -NamedValues $RecordNamedValues -SMSServerConnection $SMSConnection)
        $SMSCommands = $SMSCommands + $SMSCommand
        If ($PSCmdlet.ShouldProcess("$Item","Remove Access Code $AccessCodeID")) {
            $SMSCommand.Execute()
        }
        
        if ($ReturnSMSCommand) {
            return $SMSCommand
        }
	}
	End {
        if ($Wait) {
            Sleep (($SMSConnection.SMSImportIntervalMinutes * 60) + 1)
        }
	}
}

function Get-SMSAccessCode {
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="AccessCodeID to find.")]
        [int]$AccessCodeID,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Limit results to specific CompanyID")]
        [int]$CompanyID,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Find access code(s) by name, SQL wildcards allowed.")]
        [Alias("AccessGroupName")]
        [string]$AccessCodeName,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {
        [String] $SQLCommand = "Select * from AccessCodeTable"

        [String]$WHERE = ""

        if ($AccessCodeID) {
            $WHERE = $WHERE + "AccessCodeID = $AccessCodeID"
        }

        If ($CompanyID) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "CompanyID = $CompanyID"
        }

        If ($AccessCodeName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "AccessCodeName Like '$AccessCodeName'"
        }

        if ($WHERE) {
            $SQLCommand = $SQLCommand + " WHERE " + $WHERE
        }

        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue
	}
}

function Get-SMSAccessRights {
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="CardID to find.")]
        [int]$CardID,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Find access code(s) by name, SQL wildcards allowed.")]
        [alias("AccessCodeName")]
        [string]$AccessGroupName,
        [Parameter(Mandatory=$false,HelpMessage="Show only those enabled, defaults to True")]
        [switch]$IsEnabled = $true,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$false,HelpMessage="Return all fields")]
        [switch] $Extended,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {
        [String] $SQLCommand = ""
        if ($Extended) {
            $SQLCommand = "Select * from ViewAccessRights"
        } else {
            $SQLCommand = "Select CardID, AccessGroupName from ViewAccessRights"
        }

        [String]$WHERE = ""

        If ($CardID) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "CardID = $CardID"
        }

        If ($AccessGroupName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "AccessGroupName Like '$AccessGroupName'"
        }

        If ($IsEnabled) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "IsEnabled = '$IsEnabled'"
        }

        if ($WHERE) {
            $SQLCommand = $SQLCommand + " WHERE " + $WHERE
        }

        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAX" -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAX" -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue
	}
}

function Get-SMSCard {
    #This is a non-standard SMS cmdlet as it directly accesses the database, use at your own risk
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
		[Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$true,HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipelineByPropertyName=$true,HelpMessage="Employee reference number, typically the employee number or employee ID.  If passed ADUser object, uses EmployeeID")]
        [alias("EmployeeID")]
		[String]$EmployeeReference,
        [Parameter(Mandatory=$false,Position=2,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$false,Position=3,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$false,Position=4,ValueFromPipelineByPropertyName=$true,HelpMessage="Specify CustomerCode")]
        [Alias("CustomerCode")]
        [String]$CustomerCodeNumber,
        [Parameter(Mandatory=$false)]
		[int]$CardID,
        [Parameter(Mandatory=$false,HelpMessage="Returns inactive cards too, default is to only return active cards.")]
        [Switch]$IncludeInactive,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$false,HelpMessage="Return all fields")]
        [switch]$Extended,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {

        [String] $SQLCommand = ""
        if ($Extended) {
            $SQLCommand = "Select * from ViewSMSCardHolders"
        } else {
            $SQLCommand = "Select CardID, FirstName, LastName, InitLet, EmployeeNumber, Visitor, CompanyID, CustomerCodeNumber, CardNumber, PrimaryCard, PINNumber, Inactive, ActiveDateTime, InactiveDateTime, ExpirationDateTime from ViewSMSCardHolders"
        }

        [String]$WHERE = ""

        #Build WHERE cluases
        If ($CardID) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "CardID = $CardID"
        }

        If ($CardNumber) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "CardNumber = $CardNumber"
        }

        If ($EmployeeReference) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $EmployeeReference = $EmployeeReference.Replace("*","%")
            $WHERE = $WHERE + "EmployeeNumber Like '$EmployeeReference'"
        }
        
        If ($CustomerCodeNumber) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "CustomerCodeNumber = '$CustomerCodeNumber'"
        }

        If ($LastName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $LastName = $LastName.Replace("*","%")
            $WHERE = $WHERE + "LastName Like '$LastName'"
        }
        
        If ($FirstName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $FirstName = $FirstName.Replace("*","%")
            $WHERE = $WHERE + "FirstName Like '$FirstName'"
        }

        if ($IncludeInactive) {
            #Do not filter it
        }else {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "Inactive = 0"
        }
        
        if ($WHERE) {
            $SQLCommand = $SQLCommand + " WHERE " + $WHERE
        }

        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAX" -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAX" -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue
	}
}

function Get-SMSAlarms {
    #This is a non-standard SMS cmdlet as it directly accesses the database, use at your own risk
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
		[Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$true,HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipelineByPropertyName=$true,HelpMessage="WhereName or Location description")]
        [alias("Location")]
		[String]$WhereName,
        [Parameter(Mandatory=$false,Position=2,ValueFromPipelineByPropertyName=$true,HelpMessage="Condition description of transaction.")]
        [alias("Condition")]
		[String]$TxnConditionName,
        [Parameter(Mandatory=$false,Position=3,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$false,Position=4,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$false,Position=5,ValueFromPipelineByPropertyName=$false,HelpMessage="Return all fields")]
        [switch] $Extended,
        [Parameter(Mandatory=$false,Position=6,ValueFromPipelineByPropertyName=$false,HelpMessage="Return top number of transactions, defaults to 500")]
        [int] $Top = 500,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {

        [String] $SQLCommand = ""
        if ($Extended) {
            $SQLCommand = "Select TOP $Top * from ViewAlarmEventTransaction "
        } else {
            $SQLCommand = "Select TOP $Top TxnID, DateTimeofTxn, CompanyID, CompanyName, WhereName, TxnConditionName, AlarmPriority, AlarmColour, AlarmInstructionText, FirstName, LastName, CustomerCodeNumber, CardNumber from ViewAlarmEventTransaction "
        }

        [String]$WHERE = ""

        #Build WHERE cluases
        If ($CardNumber) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "CardNumber = $CardNumber"
        }

        If ($WhereName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WhereName = $WhereName.Replace("*","%")
            $WHERE = $WHERE + "WhereName Like '$WhereName'"
        }

        If ($TxnConditionName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $TxnConditionName = $TxnConditionName.Replace("*","%")
            $WHERE = $WHERE + "TxnConditionName Like '$TxnConditionName'"
        }
        
        If ($LastName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $LastName = $LastName.Replace("*","%")
            $WHERE = $WHERE + "LastName Like '$LastName'"
        }
        
        If ($FirstName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $FirstName = $FirstName.Replace("*","%")
            $WHERE = $WHERE + "FirstName Like '$FirstName'"
        }
        
        if ($WHERE) {
            $SQLCommand = $SQLCommand + " WHERE " + $WHERE
        }

        $SQLCommand = $SQLCommand + " ORDER BY DateTimeOfTxn DESC"

        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAXTxn" -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAXTxn" -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue
	}
}

function Get-SMSActivity {
    #This is a non-standard SMS cmdlet as it directly accesses the database, use at your own risk
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
		[Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$true,HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipelineByPropertyName=$true,HelpMessage="WhereName or Location description")]
        [alias("Location")]
		[String]$WhereName,
        [Parameter(Mandatory=$false,Position=2,ValueFromPipelineByPropertyName=$true,HelpMessage="Condition description of transaction.")]
        [alias("Condition")]
		[String]$TxnConditionName,
        [Parameter(Mandatory=$false,Position=3,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$false,Position=4,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$false,Position=5,ValueFromPipelineByPropertyName=$false,HelpMessage="Return all fields")]
        [switch] $Extended,
        [Parameter(Mandatory=$false,Position=6,ValueFromPipelineByPropertyName=$false,HelpMessage="Return top number of transactions, defaults to 500")]
        [int] $Top = 500,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {

        [String] $SQLCommand = ""
        if ($Extended) {
            $SQLCommand = "Select TOP $Top * from ActivityDataView "
        } else {
            $SQLCommand = "Select TOP $Top TxnID, DateTimeofTxn, CompanyID, CompanyName, WhereName, TxnConditionName, AlarmPriority, AlarmColour, AlarmInstructionText, FirstName, LastName, CustomerCodeNumber, CardNumber from ActivityDataView "
        }

        [String]$WHERE = ""

        #Build WHERE cluases
        If ($CardNumber) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "CardNumber = $CardNumber"
        }

        If ($WhereName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WhereName = $WhereName.Replace("*","%")
            $WHERE = $WHERE + "WhereName Like '$WhereName'"
        }

        If ($TxnConditionName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $TxnConditionName = $TxnConditionName.Replace("*","%")
            $WHERE = $WHERE + "TxnConditionName Like '$TxnConditionName'"
        }
        
        If ($LastName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $LastName = $LastName.Replace("*","%")
            $WHERE = $WHERE + "LastName Like '$LastName'"
        }
        
        If ($FirstName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $FirstName = $FirstName.Replace("*","%")
            $WHERE = $WHERE + "FirstName Like '$FirstName'"
        }
        
        if ($WHERE) {
            $SQLCommand = $SQLCommand + " WHERE " + $WHERE
        }
        
        $SQLCommand = $SQLCommand + " ORDER BY DateTimeOfTxn DESC"

        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAXTxn" -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAXTxn" -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue
	}
}


function Get-SMSCardLocation {
    #This is a non-standard SMS cmdlet as it directly accesses the database, use at your own risk
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
		[Parameter(Mandatory=$false,Position=0,ValueFromPipeline=$true,HelpMessage="Card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$false,Position=1,ValueFromPipelineByPropertyName=$true,HelpMessage="WhereName or Location description")]
        [alias("Location")]
		[String]$LastTxn,
        [Parameter(Mandatory=$false,Position=2,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders Last name")]
        [alias("Surname")]
		[String]$LastName,
        [Parameter(Mandatory=$false,Position=3,ValueFromPipelineByPropertyName=$true,HelpMessage="Card holders First name")]
        [alias("GivenName")]
		[String]$FirstName,
        [Parameter(Mandatory=$false,Position=4,ValueFromPipelineByPropertyName=$false,HelpMessage="Return all fields")]
        [switch] $Extended,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {

        [String] $SQLCommand = ""
        if ($Extended) {
            $SQLCommand = "Select * from LocatorByCardInfoView "
        } else {
            $SQLCommand = "Select CompanyID, FirstName, LastName, InitLet, CardNumber, CustomerCodeNumber, LastTxn, LastTxnDateTime from LocatorByCardInfoView "
        }

        [String]$WHERE = ""

        #Build WHERE cluases
        If ($CardNumber) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $WHERE = $WHERE + "CardNumber = $CardNumber"
        }

        If ($LastTxn) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $LastTxn = $LastTxn.Replace("*","%")
            $WHERE = $WHERE + "LastTxn Like '$LastTxn'"
        }
      
        If ($LastName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $LastName = $LastName.Replace("*","%")
            $WHERE = $WHERE + "LastName Like '$LastName'"
        }
        
        If ($FirstName) {
            if ($WHERE) {
                $WHERE = $WHERE + " AND "
            }
            $FirstName = $FirstName.Replace("*","%")
            $WHERE = $WHERE + "FirstName Like '$FirstName'"
        }
        
        if ($WHERE) {
            $SQLCommand = $SQLCommand + " WHERE " + $WHERE 
        }

        $SQLCommand = $SQLCommand + " ORDER BY LastTxnDateTime DESC"

        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAX" -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAX" -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue
	}
}

#Note, this is significantly limited in what it can copy as it uses Add-SMSCard
function Copy-SMSCard {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage="Find Card number to copy")]
		[int]$CopyCardNumber,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="New card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$false,HelpMessage="Find Customer/Facility Code to match card number against.  If not provided, uses CustomerCode")]
        [int]$CopyCustomerCode,
        [Parameter(Mandatory=$true,HelpMessage="New Customer/Facility Code to set.  If unchanged, specify this parameter and CopyCustomerCode is not required.")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {
        if (!$CopyCustomerCode) {
            $CopyCustomerCode = $CustomerCode
        }
        #New card issued under new CardID
        $CopyCard = Get-SMSCard -CardNumber $CopyCardNumber -CustomerCode $CopyCustomerCode -Extended -SMSConnection $SMSConnection
        
        #Get all OtherNamedValues
        $OtherNamedValues = @{}
        forEach ($property in (Get-Member -InputObject $CopyCard | Where -Property "MemberType" -EQ -Value "Property").Name) {
            switch ($property) {
                "CompanyID" {$OtherNamedValues.Add($property.Replace(" ", ""), $CopyCard.Item($property))}
                #"CardIssueLevel" {$OtherNamedValues.Add($property.Replace(" ", ""), $CopyCard.Item($property))}  #Only if turned on in system, TODO, add check
                "PIN" {$OtherNamedValues.Add($property.Replace(" ", ""), $CopyCard.Item($property))}
                "DeactivateAtThreatLevel" {$OtherNamedValues.Add($property.Replace(" ", ""), $CopyCard.Item($property))}
                default {
                    #Handle PersonalData, ignore all the rest
                    if ($property.Replace(" ", "") -like "PersonalData*") {
                        $OtherNamedValues.Add($property.Replace(" ", ""), $CopyCard.Item($property))
                    }
                }
            }
        }
        
        #Add the new card
        $result = Add-SMSCard -CardNumber $CardNumber -CustomerCode $CustomerCode -EmployeeReference $CopyCard.EmployeeNumber -LastName $CopyCard.LastName -FirstName $CopyCard.FirstName -InitLet $CopyCard.InitLet -ActiveDate $CopyCard.ActiveDateTime -InactiveDate $CopyCard.InactiveDateTime -OtherNamedValues $OtherNamedValues -ReturnSMSCommand
        
        #Match access rights
        Get-SMSAccessRights -CardID $CopyCard.CardID | Get-SMSAccessCode | Add-SMSAccessRights -CardNumber $CardNumber -CustomerCode $CustomerCode

        #Note, currently has no ability to do other types of access rights
	}
}

function Replace-SMSCard {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Old card number to find")]
		[int]$CardNumber,
        [Parameter(Mandatory=$true,HelpMessage="New card number")]
		[int]$NewCardNumber,
        [Parameter(Mandatory=$true,ValueFromPipelineByPropertyName=$true,HelpMessage="Old Customer/Facility Code to match card number against.")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="New Customer/Facility Code to set.  If not provided, defaults to current CustomerCodeNumber.")]
        [int]$NewCustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {
        if (!$NewCustomerCode) {
            $NewCustomerCode = $CustomerCode
        }
        #Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database multiMax -Query "Update dbo.CardInfoTable SET CardNumber = $NewCardNumber, CustomerCodeNumber = $NewCustomerCodeNumber WHERE CardNumber = $CardNumber AND CustomerCodeNumber = $CustomerCode"
        Copy-SMSCard -CopyCardNumber $CardNumber -CardNumber $NewCardNumber -CopyCustomerCode $CustomerCode -CustomerCode $NewCustomerCode -SMSConnection $SMSConnection
        Disable-SMSCard -CardNumber $CardNumber -CustomerCode $CustomerCode -SMSConnection $SMSConnection
	}
}

function Get-SMSRecordsToProcess {
	[CmdletBinding(SupportsShouldProcess=$false)]
	Param(
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {
        [String] $SQLCommand = "SELECT Count([RecordStatus]) AS RecordsToProcess FROM [multiMAXImport].[dbo].[DataImportTable] WHERE [RecordStatus] = 0"

        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database $SMSConnection.SMSImportDatabase -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue.RecordsToProcess
	}
}

function Get-SMSGroupDoorPermission {
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


function Get-SMSGroupDoorAccessCode {
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
            $Groups = Get-SMSGroupDoorPermission -Identity $Identity -OU $OU
        } else {
            $Groups = Get-SMSGroupDoorPermission -Identity $Identity -ADGroupPrefix $ADGroupPrefix
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

function Sync-SMSwithAD {
	[CmdletBinding(SupportsShouldProcess=$true,DefaultParameterSetName="ByADUser")]
	Param(
		[Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="ByADUser",HelpMessage="Specific ADUser to sync")]
		[Microsoft.ActiveDirectory.Management.ADUser[]]$ADUser,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="ByADGroup",HelpMessage="Specific ADGroup to sync.  All users of the group are processed if they have and EmployeeID")]
		[Microsoft.ActiveDirectory.Management.ADGroup[]]$ADGroup,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,ParameterSetName="ByDoor",HelpMessage="To sync by door AD security groups, specify an OU to find groups in.")]
        [String]$OU,
        [Parameter(Mandatory=$false,HelpMessage="Prefix of AD group name to remove when looking up Access codes")]
        [String]$ADGroupPrefix,
        [Parameter(Mandatory=$true,HelpMessage="Customer/Facility Code to use.")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {
        If ($OU) {   #By DOOR ADGroup
            $Doors = Get-ADGroup -SearchBase $OU -Filter "Name -like '$($ADGroupPrefix)*'"
            forEach ($Door in $Doors) {
                #for a specific door, determine what users should have access
                $usersShould = ($Door | Get-ADGroupMember -Recursive | Get-ADUser -Properties EmployeeID | %{Get-SMSCard -EmployeeReference $_.EmployeeID}).CardID


                #for a specifc door, determine who currently has access
                $usersDo = (Get-SMSAccessRights -AccessGroupName ($Door.Name.Replace($ADGroupPrefix, ""))).CardID

                #TODO Need to handle operators who are not IN AD yet
                $results = Compare-Object -ReferenceObject $usersDo -DifferenceObject $usersShould

                $toRemove = ($results | Where "SideIndicator" -EQ "=>").InputObject
                $toAdd = ($results | Where "SideIndicator" -EQ "<=").InputObject

                $accessCodeID = (Get-SMSAccessCode -AccessCodeName ($Door.Name.Replace($ADGroupPrefix, ""))).AccessCodeID
                forEach ($result in $toAdd) {
                    #Write-Verbose "Adding AccessCodeID $accessCodeID to card $result"
                    Add-SMSAccessRights -CardNumber ((Get-SMSCard -CardID $result).CardNumber) -AccessCodeID $accessCodeID -CustomerCode $CustomerCode -SMSConnection $SMSConnection
                }

                forEach ($result in $toRemove) {
                    #Write-Verbose "Removing AccessCodeID $accessCodeID from card $result"
                    Remove-SMSAccessRights -CardNumber ((Get-SMSCard -CardID $result).CardNumber) -AccessCodeID $accessCodeID -CustomerCode $CustomerCode -SMSConnection $SMSConnection
                }
            }
        } else {     #By ADGroup or ADUser

            If ($ADGroup) {
                $ADUser = $ADGroup | Get-ADGroupMember -Recursive
            }

            forEach ($SpecificUser in $ADUser) {
            
                $User = $SpecificUser | Get-ADUser -Properties EmployeeID
                if ($User.EmployeeID) {

                    #Get what it should be
                    $adcodes = Get-GroupMembershipRecursive $User | where name -Like "$ADGroupPrefix*" | select Name | %{Get-SMSAccessCode -AccessCodeName ($_.Name.Replace("$ADGroupPrefix", "")) -SMSConnection $SMSConnection}

                    #Get what it is
                    $card = Get-SMSCard -EmployeeReference ($User.EmployeeID) -SMSConnection $SMSConnection
                    $smscardcodes = Get-SMSAccessRights -CardID ($card.CardID) -Extended -SMSConnection $SMSConnection
                    if (!$smscardcodes) {
                        $smscodes = @()
                    } else {
                        $smscodes = $smscardcodes.AccessGroupName | %{Get-SMSAccessCode -AccessCodeName $_ -SMSConnection $SMSConnection}
                    }

                    #compare
                    $results = Compare-Object -ReferenceObject $smscodes -DifferenceObject $adcodes -Property AccessCodeID
                    $toAdd = $results | Where "SideIndicator" -EQ "=>"
                    $toRemove = $results | Where "SideIndicator" -EQ "<="

                    forEach ($result in $toAdd) {
                        Add-SMSAccessRights -CardNumber ($card.CardNumber) -AccessCodeID ($result.AccessCodeID) -CustomerCode $CustomerCode -SMSConnection $SMSConnection
                    }

                    forEach ($result in $toRemove) {
                        Remove-SMSAccessRights -CardNumber ($card.CardNumber) -AccessCodeID ($result.AccessCodeID) -CustomerCode $CustomerCode -SMSConnection $SMSConnection
                    }
                } else {
                    Write-Warning "Skipping $($SpecificUser.SamAccountName) as it is missing an EmployeeID"
                }
            }
        }
	}
}

Export-ModuleMember -Function *

#Export-ModuleMember Get-SMSServerConnection, Disable-SMSCard, Enable-SMSCard, Set-SMSCard, Add-SMSCard, Remove-SMSCard, Add-SMSAccessRights, Remove-SMSAccessRights, Get-SMSAccessCode, Get-SMSCard, Get-SMSAlarms, Get-SMSCardLocation, Get-SMSActivity, Get-SMSAccessRights, Copy-SMSCard, Replace-SMSCard, Get-SMSRecordsToProcess