# SAP.ps1 - SAP

$Log_MaskableKeys = @(
    
)

# System functions
#
function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log info "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"

    if ($Connection) {
        @(
            @{
                name = 'Hostname'
                type = 'textbox'
                label = 'Hostname'
                description = 'The Hostname of the SAP server where the SAP system in installed'
                value = ''
                required = $true
            }
            @{
                name = 'SysId'
                type = 'textbox'
                label = 'SysId'
                description = 'The ID of the SAP system'
                value = ''
                required = $true
            }
            @{
                name = 'SysNr'
                type = 'textbox'
                label = 'SysNr'
                description = 'The Number of SAP System'
                value = ''
                required = $true
            }
            @{
                name = 'Client'
                type = 'textbox'
                label = 'Client'
                description = 'The Client of the SAP system'
                value = ''
                required = $true
            }
            @{
                name = 'Username'
                type = 'textbox'
                label = 'Username'
                description = 'The Username to connect the SAP system'
                value = ''
                required = $true
            }
            @{
                name = 'Password'
                type = 'textbox'
                password = $true
                label = 'Password'
                description = 'The Password to connect the SAP system'
                value = ''
                required = $true
            }
            @{
                name = 'SAPdllDirectoryPath'
                type = 'textbox'
                label = 'SAP DLL Path'
                description = 'The directory where the SAP connector dll''s are located'
                value = 'C:\Program Files\SAP\SAP_DotNetConnector3_Net40_x64'
                required = $true
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                description = ''
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                description = ''
                value = 1
            }
        )
    }

    if ($TestConnection) {
        # Test is a failure only if an exception is thrown.
        $response = Idm-UsaStatesRead -SystemParams $ConnectionParams
    }

    if ($Configuration) {
        @()
    }

    Log info "Done"
}

function Idm-OnUnload {
}

#
# Object CRUD functions
#

$Global:Properties = @{
    SAP_User = @(
        @{ displayName = "Username"; name = "USERNAME"; options = @('default','key') }    
        @{ displayName = "account_id"; area = "LOGONDATA"; name = "ACCNT"; options = @() }
        @{ displayName = ""; area = ""; name = "ADDR_NO"; options = @() }
        @{ displayName = ""; area = ""; name = "ADR_NOTES"; options = @() }
        @{ displayName = ""; area = ""; name = "BCODE"; options = @() }
        @{ displayName = ""; area = ""; name = "BIRTH_NAME"; options = @() }
        @{ displayName = "building"; area = "ADDRESS"; name = "BUILDING"; options = @() }
        @{ displayName = ""; area = ""; name = "BUILDING_P"; options = @() }
        @{ displayName = ""; area = ""; name = "BUILD_LONG"; options = @() }
        @{ displayName = ""; area = ""; name = "CATTKENNZ"; options = @() }
        @{ displayName = ""; area = ""; name = "CHCKSTATUS"; options = @() }
        @{ displayName = "city"; area = "ADDRESS"; name = "CITY"; options = @() }
        @{ displayName = ""; area = ""; name = "CITY_NO"; options = @() }
        @{ displayName = ""; area = ""; name = "CLASS"; options = @() }
        @{ displayName = ""; area = ""; name = "CODVC"; options = @() }
        @{ displayName = ""; area = ""; name = "CODVN"; options = @() }
        @{ displayName = ""; area = ""; name = "CODVS"; options = @() }
        @{ displayName = ""; area = ""; name = "COMM_TYPE"; options = @() }
        @{ displayName = ""; area = ""; name = "COUNTRY"; options = @() }
        @{ displayName = ""; area = ""; name = "COUNTRYISO"; options = @() }
        @{ displayName = ""; area = ""; name = "COUNTY"; options = @() }
        @{ displayName = ""; area = ""; name = "COUNTY_CODE"; options = @() }
        @{ displayName = ""; area = ""; name = "C_O_NAME"; options = @() }
        @{ displayName = ""; area = ""; name = "DATFM"; options = @() }
        @{ displayName = ""; area = ""; name = "DCPFM"; options = @() }
        @{ displayName = ""; area = ""; name = "DELIV_DIS"; options = @() }
        @{ displayName = ""; area = ""; name = "DELI_SERV_NUMBER"; options = @() }
        @{ displayName = ""; area = ""; name = "DELI_SERV_TYPE"; options = @() }
        @{ displayName = "department"; area = "ADDRESS"; name = "DEPARTMENT"; options = @('default') }
        @{ displayName = ""; area = ""; name = "DISTRCT_NO"; options = @() }
        @{ displayName = ""; area = ""; name = "DISTRICT"; options = @() }
        @{ displayName = ""; area = ""; name = "DONT_USE_P"; options = @() }
        @{ displayName = ""; area = ""; name = "DONT_USE_S"; options = @() }
        @{ displayName = ""; area = ""; name = "E_MAIL"; options = @() }
        @{ displayName = "facsimile_extension"; area = "ADDRESS"; name = "FAX_EXTENS"; options = @('default') }
        @{ displayName = "facsimile_number"; area = "ADDRESS"; name = "FAX_NUMBER"; options = @('default') }
        @{ displayName = "first_name"; area = "ADDRESS"; name = "FIRSTNAME"; options = @('default') }
        @{ displayName = "floor"; area = "ADDRESS"; name = "FLOOR"; options = @() }
        @{ displayName = ""; area = ""; name = "FLOOR_P"; options = @() }
        @{ displayName = "full_name"; area = "ADDRESS"; name = "FULLNAME"; options = @('default') }
        @{ displayName = ""; area = ""; name = "FULLNAME_X"; options = @() }
        @{ displayName = "function"; area = "ADDRESS"; name = "FUNCTION"; options = @() }
        @{ displayName = ""; area = ""; name = "GLOBAL_LOCK_STATE"; options = @('default') }
        @{ displayName = "vaild_to"; area = "LOGONDATA"; name = "GLTGB"; options = @() }
        @{ displayName = "valid_from"; area = "LOGONDATA"; name = "GLTGV"; options = @() }
        @{ displayName = "scn_permit_sap_gui_checkbox"; area = "SNC"; name = "GUIFLAG"; options = @() }
        @{ displayName = ""; area = ""; name = "HOMECITYNO"; options = @() }
        @{ displayName = ""; area = ""; name = "HOME_CITY"; options = @() }
        @{ displayName = "house"; area = "ADDRESS"; name = "HOUSE_NO"; options = @() }
        @{ displayName = ""; area = ""; name = "HOUSE_NO2"; options = @() }
        @{ displayName = ""; area = ""; name = "HOUSE_NO3"; options = @() }
        @{ displayName = ""; area = ""; name = "INHOUSE_ML"; options = @() }
        @{ displayName = "initials"; area = "ADDRESS"; name = "INITIALS"; options = @() }
        @{ displayName = ""; area = ""; name = "INITS_SIG"; options = @() }
        @{ displayName = "cost_center"; area = "DEFAULTS"; name = "KOSTL"; options = @('default') }
        @{ displayName = "language"; area = "DEFAULTS"; name = "LANGU"; options = @() }
        @{ displayName = ""; area = ""; name = "LANGUCPISO"; options = @() }
        @{ displayName = ""; area = ""; name = "LANGUP_ISO"; options = @() }
        @{ displayName = ""; area = ""; name = "LANGU_CR_P"; options = @() }
        @{ displayName = ""; area = ""; name = "LANGU_ISO"; options = @() }
        @{ displayName = ""; area = ""; name = "LANGU_P"; options = @() }
        @{ displayName = "last_name"; area = "ADDRESS"; name = "LASTNAME"; options = @('default') }
        @{ displayName = ""; area = ""; name = "LOCAL_LOCK_STATE"; options = @('default') }
        @{ displayName = ""; area = ""; name = "LOCATION"; options = @() }
        @{ displayName = ""; area = ""; name = "LOCK_STATE"; options = @('default') }
        @{ displayName = ""; area = ""; name = "LTIME"; options = @() }
        @{ displayName = "middle_name"; area = "ADDRESS"; name = "MIDDLENAME"; options = @() }
        @{ displayName = ""; area = ""; name = "NAMCOUNTRY"; options = @() }
        @{ displayName = ""; area = ""; name = "NAME"; options = @() }
        @{ displayName = ""; area = ""; name = "NAMEFORMAT"; options = @() }
        @{ displayName = ""; area = ""; name = "NAME_2"; options = @() }
        @{ displayName = ""; area = ""; name = "NAME_3"; options = @() }
        @{ displayName = ""; area = ""; name = "NAME_4"; options = @() }
        @{ displayName = ""; area = ""; name = "NICKNAME"; options = @() }
        @{ displayName = ""; area = ""; name = "NO_PASSWORD_LOCK_STATE"; options = @('default') }
        @{ displayName = ""; area = ""; name = "PASSCODE"; options = @() }
        @{ displayName = ""; area = ""; name = "PBOXCIT_NO"; options = @() }
        @{ displayName = ""; area = ""; name = "PCODE1_EXT"; options = @() }
        @{ displayName = ""; area = ""; name = "PCODE2_EXT"; options = @() }
        @{ displayName = ""; area = ""; name = "PCODE3_EXT"; options = @() }
        @{ displayName = ""; area = ""; name = "PERS_NO"; options = @() }
        @{ displayName = "secure_network_communication"; area = "SNC"; name = "PNAME"; options = @() }
        @{ displayName = ""; area = ""; name = "POBOX_CTRY"; options = @() }
        @{ displayName = "postal_code"; area = "ADDRESS"; name = "POSTL_COD1"; options = @() }
        @{ displayName = ""; area = ""; name = "POSTL_COD2"; options = @() }
        @{ displayName = ""; area = ""; name = "POSTL_COD3"; options = @() }
        @{ displayName = ""; area = ""; name = "PO_BOX"; options = @() }
        @{ displayName = ""; area = ""; name = "PO_BOX_CIT"; options = @() }
        @{ displayName = ""; area = ""; name = "PO_BOX_LOBBY"; options = @() }
        @{ displayName = ""; area = ""; name = "PO_BOX_REG"; options = @() }
        @{ displayName = ""; area = ""; name = "PO_CTRYISO"; options = @() }
        @{ displayName = ""; area = ""; name = "PO_W_O_NO"; options = @() }
        @{ displayName = ""; area = ""; name = "PREFIX1"; options = @() }
        @{ displayName = ""; area = ""; name = "PREFIX2"; options = @() }
        @{ displayName = ""; area = ""; name = "PWDSALTEDHASH"; options = @() }
        @{ displayName = ""; area = ""; name = "REGIOGROUP"; options = @() }
        @{ displayName = ""; area = ""; name = "REGION"; options = @() }
        @{ displayName = "room_number"; area = "ADDRESS"; name = "ROOM_NO"; options = @() }
        @{ displayName = ""; area = ""; name = "ROOM_NO_P"; options = @() }
        @{ displayName = ""; area = ""; name = "SECONDNAME"; options = @() }
        @{ displayName = ""; area = ""; name = "SECURITY_POLICY"; options = @('default') }
        @{ displayName = ""; area = ""; name = "SORT1"; options = @() }
        @{ displayName = ""; area = ""; name = "SORT1_P"; options = @() }
        @{ displayName = ""; area = ""; name = "SORT2"; options = @() }
        @{ displayName = ""; area = ""; name = "SORT2_P"; options = @() }
        @{ displayName = "delete_after_output"; area = "DEFAULTS"; name = "SPDA"; options = @() }
        @{ displayName = "output_immediately"; area = "DEFAULTS"; name = "SPDB"; options = @() }
        @{ displayName = "spool_output_service"; area = "DEFAULTS"; name = "SPLD"; options = @() }
        @{ displayName = ""; area = ""; name = "SPLG"; options = @() }
        @{ displayName = ""; area = ""; name = "START_MENU"; options = @() }
        @{ displayName = ""; area = ""; name = "STCOD"; options = @() }
        @{ displayName = "street"; area = "ADDRESS"; name = "STREET"; options = @() }
        @{ displayName = ""; area = ""; name = "STREET_NO"; options = @() }
        @{ displayName = ""; area = ""; name = "STR_ABBR"; options = @() }
        @{ displayName = ""; area = ""; name = "STR_SUPPL1"; options = @() }
        @{ displayName = ""; area = ""; name = "STR_SUPPL2"; options = @() }
        @{ displayName = ""; area = ""; name = "STR_SUPPL3"; options = @() }
        @{ displayName = ""; area = ""; name = "TAXJURCODE"; options = @() }
        @{ displayName = "telephone_extension"; area = "ADDRESS"; name = "TEL1_EXT"; options = @() }
        @{ displayName = "telephone_number"; area = "ADDRESS"; name = "TEL1_NUMBR"; options = @() }
        @{ displayName = ""; area = ""; name = "TIMEFM"; options = @() }
        @{ displayName = ""; area = ""; name = "TIME_ZONE"; options = @() }
        @{ displayName = ""; area = ""; name = "TITLE"; options = @() }
        @{ displayName = ""; area = ""; name = "TITLE_ACA1"; options = @() }
        @{ displayName = ""; area = ""; name = "TITLE_ACA2"; options = @() }
        @{ displayName = ""; area = ""; name = "TITLE_P"; options = @() }
        @{ displayName = ""; area = ""; name = "TITLE_SPPL"; options = @() }
        @{ displayName = ""; area = ""; name = "TOWNSHIP"; options = @() }
        @{ displayName = ""; area = ""; name = "TOWNSHIP_CODE"; options = @() }
        @{ displayName = ""; area = ""; name = "TRANSPZONE"; options = @() }
        @{ displayName = ""; area = ""; name = "TZONE"; options = @() }
        @{ displayName = ""; area = ""; name = "USTYP"; options = @() }
        @{ displayName = ""; area = ""; name = "WRONG_LOGON_LOCK_STATE"; options = @('default') }
        @{ displayName = ""; area = ""; name = "XPCPT"; options = @() }
    )
}

function Idm-SAP_UsersRead {
    param (
        [switch] $GetMeta,
        [string] $SystemParams,
        [string] $FunctionParams
    )
    
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "SAP_User"
    
    if ($GetMeta) {
        #
        # Get meta data
        #
        Get-ClassMetaData -Class $Class
    } else {
        #
        # Execute function
        #
        $system_params   = ConvertFrom-Json2 $SystemParams
        $function_params = ConvertFrom-Json2 $FunctionParams

        # Setup Connection
        $Global:Connection = Open-SAPConnection -SystemParams $system_params -FunctionParams $function_params
        
        $properties = $function_params.properties

        if ($properties.length -eq 0) {
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }) | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_.displayName)) {
                    $_.displayName
                } else {
                    $_.name
                }
            }
        }
        
        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayName
        $properties = @($key) + (@($properties | Where-Object { $_ -ne $key }) | ForEach-Object { Get-FieldDisplay -Class $Class -Name $_ })
        
        try {
            Log info "Retrieving User List"
            $repository = $Global:Connection.Repository
            [SAP.Middleware.Connector.IRfcFunction]$bapiFunctionCall = $repository.CreateFunction('BAPI_USER_GETLIST')
            $bapiFunctionCall.SetValue("WITH_USERNAME","X")
            $bapiFunctionCall.SetValue("MAX_ROWS",99999)
            $bapiFunctionCall.Invoke($Global:Connection)
            [SAP.Middleware.Connector.IRfcTable]$returnData = $bapiFunctionCall.GetTable('USERLIST')

            Log info "Retrieving User Details"
            $i = $returnData.count

            foreach($item in $returnData) {
                $obj = @{
                    LOCK_STATE = $false
                    WRONG_LOGON_LOCK_STATE = $false
                    LOCAL_LOCK_STATE = $false
                    GLOBAL_LOCK_STATE = $false
                    NO_PASSWORD_LOCK_STATE = $false
                }

                [SAP.Middleware.Connector.IRfcFunction]$bapiFunctionCall2 = $repository.CreateFunction('BAPI_USER_GET_DETAIL')
                $bapiFunctionCall2.SetValue("USERNAME",$item.GetValue("USERNAME"))
                $bapiFunctionCall2.Invoke($Global:Connection)
            
                            
                #Address (Export)
                $export = $bapiFunctionCall2.GetObject('ADDRESS')
                foreach($prop in $export) {
                    $obj[(Get-FieldDisplay -Class $Class -Name $prop.Metadata.Name -Area 'ADDRESS')] = $export.GetValue($prop.Metadata.Name)
                }
            
                #SNC (Export)
                $export = $bapiFunctionCall2.GetObject('SNC')
                foreach($prop in $export) {
                    $obj[(Get-FieldDisplay -Class $Class -Name $prop.Metadata.Name -Area 'ADDRESS')] = $export.GetValue($prop.Metadata.Name)
                }
            
                #Logondata (Export)
                $export = $bapiFunctionCall2.GetObject('LOGONDATA')
                foreach($prop in $export) {
                    $obj[(Get-FieldDisplay -Class $Class -Name $prop.Metadata.Name -Area 'ADDRESS')] = $export.GetValue($prop.Metadata.Name)
                }
            
                #Defaults (Export)
                $export = $bapiFunctionCall2.GetObject('DEFAULTS')
                foreach($prop in $export) {
                    $obj[(Get-FieldDisplay -Class $Class -Name $prop.Metadata.Name -Area 'ADDRESS')] = $export.GetValue($prop.Metadata.Name)
                }
            
                #ISLOCKED (Export)
                $export = $bapiFunctionCall2.GetObject('ISLOCKED')
                if($export.GetValue('WRNG_LOGON') -eq "L") {
                    $obj.LOCK_STATE = $true
                    $obj.WRONG_LOGON_LOCK_STATE = $true
                }
            
                if($export.GetValue('LOCAL_LOCK') -eq "L") {
                    $obj.LOCK_STATE = $true
                    $obj.LOCAL_LOCK_STATE = $true
                }
                            
                if($export.GetValue('GLOB_LOCK') -eq "L") {
                    $obj.LOCK_STATE = $true
                    $obj.GLOBAL_LOCK_STATE = $true
                }
            
                if($export.GetValue('NO_USER_PW') -eq "L") {
                    $obj.NO_PASSWORD_LOCK_STATE = $true
                }
            
            
                # Activity Groups (Table)
                #table = [SAP.Middleware.Connector.IRfcTable]$table = $bapiFunctionCall2.GetTable('ACTIVITYGROUPS')
                # Parameters (Table)
                #[SAP.Middleware.Connector.IRfcTable]$table = $bapiFunctionCall2.GetTable('PARAMETER')
                # UCLASS (Table)
                #[SAP.Middleware.Connector.IRfcTable]$table = $bapiFunctionCall2.GetTable('UCLASSYS')

                if(($i -= 1) % 100 -eq 0) {
                    Log debug ("$($i) remaining user details to retrieve")
                }

                ([PSCustomObject]$obj) | Select-Object $properties
            }
        }
        catch {
            Log error $_
            Write-Error $_
        }
    }

    Log info "Done"
}

function Open-SAPConnection {
    param (
        $SystemParams,
        $FunctionParams
    )
    
    try {
        Add-Type -Path "$($SystemParams.SAPdllDirectoryPath)\sapnco.dll"
        Add-Type -Path "$($SystemParams.SAPdllDirectoryPath)\sapnco_utils.dll"

        $cfgParams = New-Object SAP.Middleware.Connector.RfcConfigParameters
        $cfgParams.Add("NAME", $SystemParams.Hostname)
        $cfgParams.Add("ASHOST", $SystemParams.Hostname)
        $cfgParams.Add("SYSID", $SystemParams.SysId)
        $cfgParams.Add("SYSNR", $SystemParams.SysNr)
        $cfgParams.Add("CLIENT", $SystemParams.Client)
        $cfgParams.Add("USER", $SystemParams.Username)
        $cfgParams.Add("PASSWD", $SystemParams.Password)

        [SAP.Middleware.Connector.RfcDestinationManager]::GetDestination($cfgParams)
    } catch {
        Log error $_
        Write-Error $_
    }

    Log info "Done"
}

function Get-FieldDisplay {
    param (
        [string] $Class,
        [string] $Name,
        [string] $Area
    )
    
    $fields = $Global:Properties.$Class
    $field = $fields | ? { $_.name -eq $Name}
    $ret = $Name

    if($field.displayName.length -gt 0) {
        $ret = $field.displayName
    }
    
    $ret
}

function Get-ClassMetaData {
    param (
        [string] $Class
    )
    @(
        @{
            name = 'properties'
            type = 'grid'
            label = 'Properties'
            description = 'Selected properties'
            table = @{
                rows = @( $Global:Properties.$Class | ForEach-Object {
                    @{
                        name = $_.name
                        display_name = $_.displayName
                        area = $_.area
                        usage_hint = @( @(
                            foreach ($opt in $_.options) {
                                if ($opt -notin @('default', 'idm', 'key')) { continue }

                                if ($opt -eq 'idm') {
                                    $opt.Toupper()
                                }
                                else {
                                    $opt.Substring(0,1).Toupper() + $opt.Substring(1)
                                }
                            }
                        ) | Sort-Object) -join ' | '
                    }
                })
                settings_grid = @{
                    selection = 'multiple'
                    key_column = 'name'
                    checkbox = $true
                    filter = $true
                    columns = @(
                        @{
                            name = 'name'
                            display_name = 'Name'
                        }
                        @{
                            name = 'display_name'
                            display_name = 'Display Name'
                        }
                        @{
                            name = 'area'
                            display_name = 'Area'
                        }
                        @{
                            name = 'usage_hint'
                            display_name = 'Usage hint'
                        }
                    )
                }
            }
            value = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }
    )
}