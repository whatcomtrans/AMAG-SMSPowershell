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

        #Iterate through values
        if ($NamedValues) {
            forEach ($key in $NamedValues.Keys) {
                $Collumns += ",$key"
                switch (($NamedValues.$key).GetTypeCode()) {
                    "String" {$Values += (" ,'" + $NamedValues.$key + "'")}
                    "DateTime" {$Values += (" ,'" + ([DateTime] ($NamedValues.$key)).ToString() + "'")}
                    default {$Values += (" ,'" + $NamedValues.$key + "'")}
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
        [alias("Initials")]
		[String]$MiddleName,
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Date card is set to become active, defaults to current date.")]
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
        
        $RecordNamedValues.Add("RecordRequest",0)
        $RecordNamedValues.Add("RecordStatus",0)
        $RecordNamedValues.Add("ImportNow",1)        
        
        add-BoundParams $RecordNamedValues $PSBoundParameters @("Wait","ReturnSMSCommand","SMSConnection", "OtherNamedValues")
        if ($OtherNamedValues) {
            $RecordNamedValues = $RecordNamedValues + $OtherNamedValues
        }

        $Item = "Adding new card number $CardNumber for $FirstName $LastName."
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
            $SQLCommand = "Select TOP $Top * from ViewAlarmEventTransaction"
        } else {
            $SQLCommand = "Select TOP $Top TxnID, DateTimeofTxn, CompanyID, CompanyName, WhereName, TxnConditionName, AlarmPriority, AlarmColour, AlarmInstructionText, FirstName, LastName, CustomerCodeNumber, CardNumber from ViewAlarmEventTransaction"
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
            $SQLCommand = "Select * from LocatorByCardInfoView"
        } else {
            $SQLCommand = "Select CompanyID, FirstName, LastName, InitLet, CardNumber, CustomerCodeNumber, LastTxn, LastTxnDateTime from LocatorByCardInfoView"
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

        if (!$SMSConnection.SMSImportDatabaseUsername) {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAX" -Query $SQLCommand
        } else {
            $retvalue = Invoke-Sqlcmd -ServerInstance $SMSConnection.SMSDatabaseServer -Database "multiMAX" -Username $SMSConnection.SMSImportDatabaseUsername -Password $SMSConnection.SMSImportDatabasePassword -Query $SQLCommand
        }
        return $retvalue
	}
}

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
        $CopyCard = Get-SMSCard -CardNumber $CopyCardNumber -CustomerCode $CopyCustomerCode -Extended
        $CopyCard | Add-SMSCard -CardNumber $CardNumber -CustomerCode $CustomerCode
        Get-SMSAccessRights -CardID $CopyCard.CardID | Get-SMSAccessCode | Add-SMSAccessRights -CardNumber $CardNumber -CustomerCode $CustomerCode
	}
}

function Switch-SMSCardNumber {
	[CmdletBinding(SupportsShouldProcess=$true)]
	Param(
		[Parameter(Mandatory=$true,ValueFromPipeline=$true,HelpMessage="Old card number to find")]
		[int]$OldCardNumber,
        [Parameter(Mandatory=$true,Position=0,ValueFromPipeline=$true,HelpMessage="New card number")]
		[int]$CardNumber,
        [Parameter(Mandatory=$false,HelpMessage="Old Customer/Facility Code to match card number against.  If not provided, uses CustomerCode")]
        [int]$OldCustomerCode,
        [Parameter(Mandatory=$true,HelpMessage="New Customer/Facility Code to set.  If unchanged, specify this parameter and OldCustomerCode is not required.")]
        [int]$CustomerCode,
        [Parameter(Mandatory=$false,HelpMessage="SMSConnection object, use Get-SMSServerConnection to create the object.")]
        [object]$SMSConnection=$DefaultSMSServerConnection
	)
	Process {
        if (!$OldCustomerCode) {
            $OldCustomerCode = $CustomerCode
        }
        #New card issued under new CardID
        Copy-SMSCard -CopyCardNumber $OldCardNumber -CopyCustomerCode $OldCustomerCode -CardNumber $CardNumber -CustomerCode $CustomerCode
        Remove-SMSCard -CardNumber $OldCardNumber -CustomerCode $OldCustomerCode
	}
}

Export-ModuleMember Get-SMSServerConnection, Disable-SMSCard, Enable-SMSCard, Set-SMSCard, Add-SMSCard, Remove-SMSCard, Add-SMSAccessRights, Remove-SMSAccessRights, Get-SMSAccessCode, Get-SMSCard, Get-SMSAlarms, Get-SMSCardLocation, Get-SMSAccessRights, Switch-SMSCardNumber, Copy-SMSCard