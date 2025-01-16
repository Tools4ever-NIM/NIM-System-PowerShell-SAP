# SAP.ps1 - SAP

$Log_MaskableKeys = @(
    "password"
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
    UserHT = [System.Collections.ArrayList]@()
    User = @(
        @{ displayName = 'Username'; area = 'ADDRESS'; name='USERNAME'; options = @('default','key') }
        @{ displayName = 'account_id'; area = 'LOGONDATA'; name='ACCNT'; options = @('default') }
        @{ displayName = 'ADDR_NO'; area = 'ADDRESS'; name='ADDR_NO'; options = @() }
        @{ displayName = 'ADR_NOTES'; area = 'ADDRESS'; name='ADR_NOTES'; options = @() }
        @{ displayName = 'BCODE'; area = 'LOGONDATA'; name='BCODE'; options = @() }
        @{ displayName = 'BIRTH_NAME'; area = 'ADDRESS'; name='BIRTH_NAME'; options = @() }
        @{ displayName = 'building'; area = 'ADDRESS'; name='BUILDING'; options = @('default') }
        @{ displayName = 'BUILDING_P'; area = 'ADDRESS'; name='BUILDING_P'; options = @() }
        @{ displayName = 'BUILD_LONG'; area = 'ADDRESS'; name='BUILD_LONG'; options = @() }
        @{ displayName = 'CATTKENNZ'; area = 'DEFAULTS'; name='CATTKENNZ'; options = @() }
        @{ displayName = 'CHCKSTATUS'; area = 'ADDRESS'; name='CHCKSTATUS'; options = @() }
        @{ displayName = 'city'; area = 'ADDRESS'; name='CITY'; options = @('default') }
        @{ displayName = 'CITY_NO'; area = 'ADDRESS'; name='CITY_NO'; options = @() }
        @{ displayName = 'CLASS'; area = 'LOGONDATA'; name='CLASS'; options = @() }
        @{ displayName = 'CODVC'; area = 'LOGONDATA'; name='CODVC'; options = @() }
        @{ displayName = 'CODVN'; area = 'LOGONDATA'; name='CODVN'; options = @() }
        @{ displayName = 'CODVS'; area = 'LOGONDATA'; name='CODVS'; options = @() }
        @{ displayName = 'COMM_TYPE'; area = 'ADDRESS'; name='COMM_TYPE'; options = @() }
        @{ displayName = 'COUNTRY'; area = 'ADDRESS'; name='COUNTRY'; options = @() }
        @{ displayName = 'COUNTRYISO'; area = 'ADDRESS'; name='COUNTRYISO'; options = @() }
        @{ displayName = 'COUNTY'; area = 'ADDRESS'; name='COUNTY'; options = @() }
        @{ displayName = 'COUNTY_CODE'; area = 'ADDRESS'; name='COUNTY_CODE'; options = @() }
        @{ displayName = 'C_O_NAME'; area = 'ADDRESS'; name='C_O_NAME'; options = @() }
        @{ displayName = 'DATFM'; area = 'DEFAULTS'; name='DATFM'; options = @() }
        @{ displayName = 'DCPFM'; area = 'DEFAULTS'; name='DCPFM'; options = @() }
        @{ displayName = 'DELIV_DIS'; area = 'ADDRESS'; name='DELIV_DIS'; options = @() }
        @{ displayName = 'DELI_SERV_NUMBER'; area = 'ADDRESS'; name='DELI_SERV_NUMBER'; options = @() }
        @{ displayName = 'DELI_SERV_TYPE'; area = 'ADDRESS'; name='DELI_SERV_TYPE'; options = @() }
        @{ displayName = 'department'; area = 'ADDRESS'; name='DEPARTMENT'; options = @('default') }
        @{ displayName = 'DISTRCT_NO'; area = 'ADDRESS'; name='DISTRCT_NO'; options = @() }
        @{ displayName = 'DISTRICT'; area = 'ADDRESS'; name='DISTRICT'; options = @() }
        @{ displayName = 'DONT_USE_P'; area = 'ADDRESS'; name='DONT_USE_P'; options = @() }
        @{ displayName = 'DONT_USE_S'; area = 'ADDRESS'; name='DONT_USE_S'; options = @() }
        @{ displayName = 'E_MAIL'; area = 'ADDRESS'; name='E_MAIL'; options = @() }
        @{ displayName = 'facsimile_extension'; area = 'ADDRESS'; name='FAX_EXTENS'; options = @('default') }
        @{ displayName = 'facsimile_number'; area = 'ADDRESS'; name='FAX_NUMBER'; options = @('default') }
        @{ displayName = 'first_name'; area = 'ADDRESS'; name='FIRSTNAME'; options = @('default') }
        @{ displayName = 'floor'; area = 'ADDRESS'; name='FLOOR'; options = @() }
        @{ displayName = 'FLOOR_P'; area = 'ADDRESS'; name='FLOOR_P'; options = @() }
        @{ displayName = 'full_name'; area = 'ADDRESS'; name='FULLNAME'; options = @('default') }
        @{ displayName = 'FULLNAME_X'; area = 'ADDRESS'; name='FULLNAME_X'; options = @() }
        @{ displayName = 'function'; area = 'ADDRESS'; name='FUNCTION'; options = @() }
        @{ displayName = 'GLOBAL_LOCK_STATE'; area = 'ISLOCKED'; name='GLOBAL_LOCK_STATE'; options = @('default') }
        @{ displayName = 'vaild_to'; area = 'LOGONDATA'; name='GLTGB'; options = @('default') }
        @{ displayName = 'valid_from'; area = 'LOGONDATA'; name='GLTGV'; options = @('default') }
        @{ displayName = 'scn_permit_sap_gui_checkbox'; area = 'SNC'; name='GUIFLAG'; options = @('default') }
        @{ displayName = 'HOMECITYNO'; area = 'ADDRESS'; name='HOMECITYNO'; options = @() }
        @{ displayName = 'HOME_CITY'; area = 'ADDRESS'; name='HOME_CITY'; options = @() }
        @{ displayName = 'house'; area = 'ADDRESS'; name='HOUSE_NO'; options = @('default') }
        @{ displayName = 'HOUSE_NO2'; area = 'ADDRESS'; name='HOUSE_NO2'; options = @() }
        @{ displayName = 'HOUSE_NO3'; area = 'ADDRESS'; name='HOUSE_NO3'; options = @() }
        @{ displayName = 'INHOUSE_ML'; area = 'ADDRESS'; name='INHOUSE_ML'; options = @() }
        @{ displayName = 'initials'; area = 'ADDRESS'; name='INITIALS'; options = @('default') }
        @{ displayName = 'INITS_SIG'; area = 'ADDRESS'; name='INITS_SIG'; options = @() }
        @{ displayName = 'cost_center'; area = 'DEFAULTS'; name='KOSTL'; options = @('default') }
        @{ displayName = 'language'; area = 'ADDRESS'; name='LANGU'; options = @('default') }
        @{ displayName = 'LANGUCPISO'; area = 'ADDRESS'; name='LANGUCPISO'; options = @() }
        @{ displayName = 'LANGUP_ISO'; area = 'ADDRESS'; name='LANGUP_ISO'; options = @() }
        @{ displayName = 'LANGU_CR_P'; area = 'ADDRESS'; name='LANGU_CR_P'; options = @() }
        @{ displayName = 'LANGU_ISO'; area = 'ADDRESS'; name='LANGU_ISO'; options = @() }
        @{ displayName = 'LANGU_P'; area = 'ADDRESS'; name='LANGU_P'; options = @() }
        @{ displayName = 'last_name'; area = 'ADDRESS'; name='LASTNAME'; options = @('default') }
        @{ displayName = 'LOCAL_LOCK_STATE'; area = 'ISLOCKED'; name='LOCAL_LOCK_STATE'; options = @('default') }
        @{ displayName = 'LOCATION'; area = 'ADDRESS'; name='LOCATION'; options = @() }
        @{ displayName = 'LOCK_STATE'; area = 'ISLOCKED'; name='LOCK_STATE'; options = @('default') }
        @{ displayName = 'LTIME'; area = 'LOGONDATA'; name='LTIME'; options = @() }
        @{ displayName = 'middle_name'; area = 'ADDRESS'; name='MIDDLENAME'; options = @('default') }
        @{ displayName = 'NAMCOUNTRY'; area = 'ADDRESS'; name='NAMCOUNTRY'; options = @() }
        @{ displayName = 'NAME'; area = 'ADDRESS'; name='NAME'; options = @() }
        @{ displayName = 'NAMEFORMAT'; area = 'ADDRESS'; name='NAMEFORMAT'; options = @() }
        @{ displayName = 'NAME_2'; area = 'ADDRESS'; name='NAME_2'; options = @() }
        @{ displayName = 'NAME_3'; area = 'ADDRESS'; name='NAME_3'; options = @() }
        @{ displayName = 'NAME_4'; area = 'ADDRESS'; name='NAME_4'; options = @() }
        @{ displayName = 'NICKNAME'; area = 'ADDRESS'; name='NICKNAME'; options = @() }
        @{ displayName = 'NO_PASSWORD_LOCK_STATE'; area = 'ISLOCKED'; name='NO_PASSWORD_LOCK_STATE'; options = @('default') }
        @{ displayName = 'PASSCODE'; area = 'LOGONDATA'; name='PASSCODE'; options = @() }
        @{ displayName = 'PBOXCIT_NO'; area = 'ADDRESS'; name='PBOXCIT_NO'; options = @() }
        @{ displayName = 'PCODE1_EXT'; area = 'ADDRESS'; name='PCODE1_EXT'; options = @() }
        @{ displayName = 'PCODE2_EXT'; area = 'ADDRESS'; name='PCODE2_EXT'; options = @() }
        @{ displayName = 'PCODE3_EXT'; area = 'ADDRESS'; name='PCODE3_EXT'; options = @() }
        @{ displayName = 'PERS_NO'; area = 'ADDRESS'; name='PERS_NO'; options = @() }
        @{ displayName = 'secure_network_communication'; area = 'SNC'; name='PNAME'; options = @('default') }
        @{ displayName = 'POBOX_CTRY'; area = 'ADDRESS'; name='POBOX_CTRY'; options = @() }
        @{ displayName = 'postal_code'; area = 'ADDRESS'; name='POSTL_COD1'; options = @('default') }
        @{ displayName = 'POSTL_COD2'; area = 'ADDRESS'; name='POSTL_COD2'; options = @() }
        @{ displayName = 'POSTL_COD3'; area = 'ADDRESS'; name='POSTL_COD3'; options = @() }
        @{ displayName = 'PO_BOX'; area = 'ADDRESS'; name='PO_BOX'; options = @() }
        @{ displayName = 'PO_BOX_CIT'; area = 'ADDRESS'; name='PO_BOX_CIT'; options = @() }
        @{ displayName = 'PO_BOX_LOBBY'; area = 'ADDRESS'; name='PO_BOX_LOBBY'; options = @() }
        @{ displayName = 'PO_BOX_REG'; area = 'ADDRESS'; name='PO_BOX_REG'; options = @() }
        @{ displayName = 'PO_CTRYISO'; area = 'ADDRESS'; name='PO_CTRYISO'; options = @() }
        @{ displayName = 'PO_W_O_NO'; area = 'ADDRESS'; name='PO_W_O_NO'; options = @() }
        @{ displayName = 'PREFIX1'; area = 'ADDRESS'; name='PREFIX1'; options = @() }
        @{ displayName = 'PREFIX2'; area = 'ADDRESS'; name='PREFIX2'; options = @() }
        @{ displayName = 'PWDSALTEDHASH'; area = 'LOGONDATA'; name='PWDSALTEDHASH'; options = @() }
        @{ displayName = 'REGIOGROUP'; area = 'ADDRESS'; name='REGIOGROUP'; options = @() }
        @{ displayName = 'REGION'; area = 'ADDRESS'; name='REGION'; options = @() }
        @{ displayName = 'room_number'; area = 'ADDRESS'; name='ROOM_NO'; options = @('default') }
        @{ displayName = 'ROOM_NO_P'; area = 'ADDRESS'; name='ROOM_NO_P'; options = @() }
        @{ displayName = 'SECONDNAME'; area = 'ADDRESS'; name='SECONDNAME'; options = @() }
        @{ displayName = 'SECURITY_POLICY'; area = 'LOGONDATA'; name='SECURITY_POLICY'; options = @('default') }
        @{ displayName = 'SORT1'; area = 'ADDRESS'; name='SORT1'; options = @() }
        @{ displayName = 'SORT1_P'; area = 'ADDRESS'; name='SORT1_P'; options = @() }
        @{ displayName = 'SORT2'; area = 'ADDRESS'; name='SORT2'; options = @() }
        @{ displayName = 'SORT2_P'; area = 'ADDRESS'; name='SORT2_P'; options = @() }
        @{ displayName = 'delete_after_output'; area = 'DEFAULTS'; name='SPDA'; options = @('default') }
        @{ displayName = 'output_immediately'; area = 'DEFAULTS'; name='SPDB'; options = @('default') }
        @{ displayName = 'spool_output_service'; area = 'DEFAULTS'; name='SPLD'; options = @('default') }
        @{ displayName = 'SPLG'; area = 'DEFAULTS'; name='SPLG'; options = @() }
        @{ displayName = 'START_MENU'; area = 'DEFAULTS'; name='START_MENU'; options = @() }
        @{ displayName = 'STCOD'; area = 'DEFAULTS'; name='STCOD'; options = @() }
        @{ displayName = 'street'; area = 'ADDRESS'; name='STREET'; options = @('default') }
        @{ displayName = 'STREET_NO'; area = 'ADDRESS'; name='STREET_NO'; options = @() }
        @{ displayName = 'STR_ABBR'; area = 'ADDRESS'; name='STR_ABBR'; options = @() }
        @{ displayName = 'STR_SUPPL1'; area = 'ADDRESS'; name='STR_SUPPL1'; options = @() }
        @{ displayName = 'STR_SUPPL2'; area = 'ADDRESS'; name='STR_SUPPL2'; options = @() }
        @{ displayName = 'STR_SUPPL3'; area = 'ADDRESS'; name='STR_SUPPL3'; options = @() }
        @{ displayName = 'TAXJURCODE'; area = 'ADDRESS'; name='TAXJURCODE'; options = @() }
        @{ displayName = 'telephone_extension'; area = 'ADDRESS'; name='TEL1_EXT'; options = @('default') }
        @{ displayName = 'telephone_number'; area = 'ADDRESS'; name='TEL1_NUMBR'; options = @('default') }
        @{ displayName = 'TIMEFM'; area = 'DEFAULTS'; name='TIMEFM'; options = @() }
        @{ displayName = 'TIME_ZONE'; area = 'ADDRESS'; name='TIME_ZONE'; options = @() }
        @{ displayName = 'TITLE'; area = 'ADDRESS'; name='TITLE'; options = @() }
        @{ displayName = 'TITLE_ACA1'; area = 'ADDRESS'; name='TITLE_ACA1'; options = @() }
        @{ displayName = 'TITLE_ACA2'; area = 'ADDRESS'; name='TITLE_ACA2'; options = @() }
        @{ displayName = 'TITLE_P'; area = 'ADDRESS'; name='TITLE_P'; options = @() }
        @{ displayName = 'TITLE_SPPL'; area = 'ADDRESS'; name='TITLE_SPPL'; options = @() }
        @{ displayName = 'TOWNSHIP'; area = 'ADDRESS'; name='TOWNSHIP'; options = @() }
        @{ displayName = 'TOWNSHIP_CODE'; area = 'ADDRESS'; name='TOWNSHIP_CODE'; options = @() }
        @{ displayName = 'TRANSPZONE'; area = 'ADDRESS'; name='TRANSPZONE'; options = @() }
        @{ displayName = 'TZONE'; area = 'LOGONDATA'; name='TZONE'; options = @() }
        @{ displayName = 'USTYP'; area = 'LOGONDATA'; name='USTYP'; options = @() }
        @{ displayName = 'WRONG_LOGON_LOCK_STATE'; area = 'ISLOCKED'; name='WRONG_LOGON_LOCK_STATE'; options = @('default') }
        @{ displayName = 'XPCPT'; area = 'ADDRESS'; name='XPCPT'; options = @() }
    )
    UserRoleHT = [System.Collections.ArrayList]@()
    UserRole = @(
        @{ displayName = 'Username'; name='USERNAME'; options = @('default','key') }
        @{ displayName = 'AGR_NAME'; name='AGR_NAME'; options = @('default') }
        @{ displayName = 'FRM_DATE'; name='FRM_DATE'; options = @('default') }
        @{ displayName = 'TO_DATE'; name='TO_DATE'; options = @('default') }
        @{ displayName = 'ORG_FLAG'; name='ORG_FLAG'; options = @('default') }
    )
    UserParameterHT = [System.Collections.ArrayList]@()
    UserParameter = @(
        @{ displayName = 'Username'; name='USERNAME'; options = @('default','key') }
        @{ displayName = 'parid'; name='parid'; options = @('default') }
        @{ displayName = 'parva'; name='parva'; options = @('default') }
        @{ displayName = 'partxt'; name='partxt'; options = @('default') }
    )
    UserProfileHT = [System.Collections.ArrayList]@()
    UserProfile = @(
        @{ displayName = 'Username'; name='USERNAME'; options = @('default','key') }
        @{ displayName = 'bapiprof'; name='bapiprof'; options = @('default') }
        @{ displayName = 'bapiptext'; name='bapiptext'; options = @('default') }
        @{ displayName = 'bapiaktps'; name='bapiaktps'; options = @('default') }
        @{ displayName = 'bapitype'; name='bapitype'; options = @('default') }
    )
    RoleHT = [System.Collections.ArrayList]@()
    Role = @(
        @{ displayName = 'AGR_NAME'; name='AGR_NAME'; options = @('default','key') }
        @{ displayName = 'FLAG_COLL'; name='FLAG_COLL'; options = @('default') }
        @{ displayName = 'TEXT'; name='TEXT'; options = @('default') }
    )
}

$Global:Properties.User | ForEach-Object { $Global:Properties.UserHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.UserHT = $Global:Properties.UserHT | Group-Object name -AsHashTable
$Global:Properties.UserRole | ForEach-Object { $Global:Properties.UserRoleHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.UserRoleHT = $Global:Properties.UserRoleHT | Group-Object name -AsHashTable
$Global:Properties.UserParameter | ForEach-Object { $Global:Properties.UserParameterHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.UserParameterHT = $Global:Properties.UserParameterHT | Group-Object name -AsHashTable
$Global:Properties.UserProfile | ForEach-Object { $Global:Properties.UserProfileHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.UserProfileHT = $Global:Properties.UserProfileHT | Group-Object name -AsHashTable
$Global:Properties.Role | ForEach-Object { $Global:Properties.RoleHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.RoleHT = $Global:Properties.RoleHT | Group-Object name -AsHashTable

$Global:User = [System.Collections.ArrayList]@()
$Global:UserRole = [System.Collections.ArrayList]@()
$Global:UserParameter = [System.Collections.ArrayList]@()
$Global:UserProfile = [System.Collections.ArrayList]@()

function Idm-UserRead {
    param (
        [switch] $GetMeta,
        [string] $SystemParams,
        [string] $FunctionParams
    )
    
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "User"
    
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
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }
        
        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        $displayProperties = $properties | ForEach-Object { $Global:Properties.UserHT[$_].displayName }

        if($Global:User.count -lt 1) {
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
                        $Global:Properties.UserHT["USERNAME"].displayName = $item.GetValue("USERNAME")
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
                        $obj[$Global:Properties.UserHT[$prop.Metadata.Name].displayName] = $export.GetValue($prop.Metadata.Name)
                    }

                    #SNC (Export)
                    $export = $bapiFunctionCall2.GetObject('SNC')
                    foreach($prop in $export) {
                        $obj[$Global:Properties.UserHT[$prop.Metadata.Name].displayName] = $export.GetValue($prop.Metadata.Name)
                    }

                    #Logondata (Export)
                    $export = $bapiFunctionCall2.GetObject('LOGONDATA')
                    foreach($prop in $export) {
                        $obj[$Global:Properties.UserHT[$prop.Metadata.Name].displayName] = $export.GetValue($prop.Metadata.Name)
                    }

                    #Defaults (Export)
                    $export = $bapiFunctionCall2.GetObject('DEFAULTS')
                    foreach($prop in $export) {
                        $obj[$Global:Properties.UserHT[$prop.Metadata.Name].displayName] = $export.GetValue($prop.Metadata.Name)
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
                    $table = [SAP.Middleware.Connector.IRfcTable]$table = $bapiFunctionCall2.GetTable('ACTIVITYGROUPS')
                    foreach($row in $table) {
                        $table_obj = @{
                            $Global:Properties.UserRoleHT["USERNAME"].displayName = $obj[$Global:Properties.UserHT["USERNAME"].displayName]
                        }
                        foreach($prop in $row) {
                            $table_obj[$prop.Metadata.Name] = $row.GetValue($prop.Metadata.Name)
                        }
                        [void]$Global:UserRole.Add([PSCustomObject]$table_obj);
                    } 

                    # Parameters (Table)
                    $table = [SAP.Middleware.Connector.IRfcTable]$table = $bapiFunctionCall2.GetTable('PARAMETER')
                    foreach($row in $table) {
                        $table_obj = @{
                            $Global:Properties.UserParameterHT["USERNAME"].displayName = $obj[$Global:Properties.UserHT["USERNAME"].displayName]
                        }
                        foreach($prop in $row) {
                            $table_obj[$prop.Metadata.Name] = $row.GetValue($prop.Metadata.Name)
                        }
                        [void]$Global:UserParameter.Add([PSCustomObject]$table_obj);
                    } 

                    # Profile (Table)
                    $table = [SAP.Middleware.Connector.IRfcTable]$table = $bapiFunctionCall2.GetTable('PROFILES')
                    foreach($row in $table) {
                        $table_obj = @{
                            $Global:Properties.UserProfileHT["USERNAME"].displayName = $obj[$Global:Properties.UserHT["USERNAME"].displayName]
                        }
                        foreach($prop in $row) {
                            $table_obj[$prop.Metadata.Name] = $row.GetValue($prop.Metadata.Name)
                        }
                        [void]$Global:UserProfile.Add([PSCustomObject]$table_obj);
                    } 

                    # UCLASS (Table)
                    <#$table = [SAP.Middleware.Connector.IRfcTable]$table = $bapiFunctionCall2.GetTable('UCLASSYS')
                    foreach($row in $table) {
                        $table_obj = @{
                            $Global:Properties.UserRoleHT["USERNAME"].displayName = $obj[$Global:Properties.UserHT["USERNAME"].displayName]
                        }
                        foreach($prop in $row) {
                            $table_obj[$prop.Metadata.Name] = $row.GetValue($prop.Metadata.Name)
                        }
                        [void]$Global:User_UCLASS.Add([PSCustomObject]$table_obj);
                    } #>
                    
                    if(($i -= 1) % 100 -eq 0) {
                        Log debug ("$($i) remaining user details to retrieve")
                    }

                    $ret = ([PSCustomObject]$obj) | Select-Object $displayProperties
                    [void]$Global:User.Add($ret);
                    
                    $ret
                }
        } else { 
            foreach($item in $Global:User) { $item }
        }
    }

    Log info "Done"
}

function Idm-UserRolesRead {
    param (
        [switch] $GetMeta,
        [string] $SystemParams,
        [string] $FunctionParams
    )
    
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "UserRole"
    
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
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }
        
        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        $displayProperties = $properties | ForEach-Object { $Global:Properties.UserRoleHT[$_].displayName }

        if($Global:User.count -lt 1) {
            Idm-UserRead -SystemParams $SystemParams | Out-Null
        } 

        foreach($item in $Global:UserRole) {            
            $item
        }
    }

    Log info "Done"
}

function Idm-UserParametersRead {
    param (
        [switch] $GetMeta,
        [string] $SystemParams,
        [string] $FunctionParams
    )
    
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "UserParameter"
    
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
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }
        
        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        $displayProperties = $properties | ForEach-Object { $Global:Properties.UserParameterHT[$_].displayName }

        if($Global:User.count -lt 1) {
            Idm-UserRead -SystemParams $SystemParams | Out-Null
        } 

        foreach($item in $Global:UserParameter) {            
            $item
        }
    }

    Log info "Done"
}

function Idm-UserProfilesRead {
    param (
        [switch] $GetMeta,
        [string] $SystemParams,
        [string] $FunctionParams
    )
    
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "UserProfile"
    
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
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }
        
        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        $displayProperties = $properties | ForEach-Object { $Global:Properties.UserProfileHT[$_].displayName }

        if($Global:User.count -lt 1) {
            Idm-UserRead -SystemParams $SystemParams | Out-Null
        } 

        foreach($item in $Global:UserProfile) {            
            $item
        }
    }

    Log info "Done"
}

function Idm-RoleRead {
    param (
        [switch] $GetMeta,
        [string] $SystemParams,
        [string] $FunctionParams
    )
    
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "Role"
    
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
            $properties = ($Global:Properties.$Class | Where-Object { $_.options.Contains('default') }).name
        }
        
        # Assure key is the first column
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).name
        $properties = @($key) + @($properties | Where-Object { $_ -ne $key })

        $displayProperties = $properties | ForEach-Object { $Global:Properties.RoleHT[$_].displayName }

        Log info "Retrieving Role List"
        $repository = $Global:Connection.Repository
        [SAP.Middleware.Connector.IRfcFunction]$bapiFunctionCall = $repository.CreateFunction('PRGN_ROLE_GETLIST')
        $bapiFunctionCall.SetValue("WITH_ROLENAME","X")
        $bapiFunctionCall.SetValue("COLL_OR_SINGLE_ONLY"," ")
        $bapiFunctionCall.SetValue("MAX_ROWS",99999)

        [SAP.Middleware.Connector.IRfcTable]$selectionRange = $bapiFunctionCall.GetTable("SELECTION_RANGE")
        $selectionRange.Append()
        $selectionRange.SetValue("SIGN", "I") 
        $selectionRange.SetValue("OPTION", "CP")    
        $selectionRange.SetValue("LOW", "*")

        $bapiFunctionCall.Invoke($Global:Connection)
        [SAP.Middleware.Connector.IRfcTable]$returnData = $bapiFunctionCall.GetTable('ROLES')
        
        foreach($row in $returnData) {
            $table_obj = @{}
            foreach($prop in $row) {
                $table_obj[$prop.Metadata.Name] = $row.GetValue($prop.Metadata.Name)
            }
            [PSCustomObject]$table_obj
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