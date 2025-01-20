# SAP.ps1 - SAP

$Log_MaskableKeys = @(
    "Password"
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
        Idm-RolesRead -SystemParams $ConnectionParams | Out-Null
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
    UserInvertHT = [System.Collections.ArrayList]@()
    User = @(
        @{ displayName = 'Username'; area = 'ADDRESS'; name='USERNAME'; options = @('default','key') }
        @{ displayName = 'account_id'; area = 'LOGONDATA'; name='ACCNT'; options = @('default','create','update') }
        @{ displayName = 'ADDR_NO'; area = 'ADDRESS'; name='ADDR_NO'; options = @() }
        @{ displayName = 'ADR_NOTES'; area = 'ADDRESS'; name='ADR_NOTES'; options = @() }
        @{ displayName = 'BCODE'; area = 'LOGONDATA'; name='BCODE'; options = @() }
        @{ displayName = 'BIRTH_NAME'; area = 'ADDRESS'; name='BIRTH_NAME'; options = @() }
        @{ displayName = 'BAPIPWD'; area = 'PASSWORD'; name='BAPIPWD'; options = @('default','create','update') }
        @{ displayName = 'BNAME_CHARGEABLE'; area = 'UCLASS'; name='BNAME_CHARGEABLE'; options = @() }
        @{ displayName = 'building'; area = 'ADDRESS'; name='BUILDING'; options = @('default') }
        @{ displayName = 'BUILDING_P'; area = 'ADDRESS'; name='BUILDING_P'; options = @('create','update') }
        @{ displayName = 'BUILD_LONG'; area = 'ADDRESS'; name='BUILD_LONG'; options = @() }
        @{ displayName = 'CATTKENNZ'; area = 'DEFAULTS'; name='CATTKENNZ'; options = @() }
        @{ displayName = 'CHCKSTATUS'; area = 'ADDRESS'; name='CHCKSTATUS'; options = @() }
        @{ displayName = 'city'; area = 'ADDRESS'; name='CITY'; options = @('default','create','update') }
        @{ displayName = 'CITY_NO'; area = 'ADDRESS'; name='CITY_NO'; options = @() }
        @{ displayName = 'CLASS'; area = 'LOGONDATA'; name='CLASS'; options = @('create','update') }
        @{ displayName = 'CLIENT'; area = 'UCLASS'; name='CLIENT'; options = @('default') }
        @{ displayName = 'CODVC'; area = 'LOGONDATA'; name='CODVC'; options = @() }
        @{ displayName = 'CODVN'; area = 'LOGONDATA'; name='CODVN'; options = @() }
        @{ displayName = 'CODVS'; area = 'LOGONDATA'; name='CODVS'; options = @() }
        @{ displayName = 'COMM_TYPE'; area = 'ADDRESS'; name='COMM_TYPE'; options = @('create','update') }
        @{ displayName = 'COUNTRY'; area = 'ADDRESS'; name='COUNTRY'; options = @() }
        @{ displayName = 'COUNTRYISO'; area = 'ADDRESS'; name='COUNTRYISO'; options = @() }
        @{ displayName = 'COUNTY'; area = 'ADDRESS'; name='COUNTY'; options = @() }
        @{ displayName = 'COUNTY_CODE'; area = 'ADDRESS'; name='COUNTY_CODE'; options = @() }
        @{ displayName = 'COUNTRY_SURCHARGE'; area = 'UCLASS'; name='COUNTRY_SURCHARGE'; options = @() }
        @{ displayName = 'C_O_NAME'; area = 'ADDRESS'; name='C_O_NAME'; options = @() }
        @{ displayName = 'DATFM'; area = 'DEFAULTS'; name='DATFM'; options = @('create','update') }
        @{ displayName = 'DCPFM'; area = 'DEFAULTS'; name='DCPFM'; options = @('create','update') }
        @{ displayName = 'DELIV_DIS'; area = 'ADDRESS'; name='DELIV_DIS'; options = @() }
        @{ displayName = 'DELI_SERV_NUMBER'; area = 'ADDRESS'; name='DELI_SERV_NUMBER'; options = @() }
        @{ displayName = 'DELI_SERV_TYPE'; area = 'ADDRESS'; name='DELI_SERV_TYPE'; options = @() }
        @{ displayName = 'department'; area = 'ADDRESS'; name='DEPARTMENT'; options = @('default','create','update') }
        @{ displayName = 'disable_password'; area = 'PASSWORD'; name='disable_password'; options = @('create','update') }
        @{ displayName = 'DISTRCT_NO'; area = 'ADDRESS'; name='DISTRCT_NO'; options = @() }
        @{ displayName = 'DISTRICT'; area = 'ADDRESS'; name='DISTRICT'; options = @() }
        @{ displayName = 'DONT_USE_P'; area = 'ADDRESS'; name='DONT_USE_P'; options = @() }
        @{ displayName = 'DONT_USE_S'; area = 'ADDRESS'; name='DONT_USE_S'; options = @() }
        @{ displayName = 'E_MAIL'; area = 'ADDRESS'; name='E_MAIL'; options = @('default','create','update') }
        @{ displayName = 'facsimile_extension'; area = 'ADDRESS'; name='FAX_EXTENS'; options = @('default','create','update') }
        @{ displayName = 'facsimile_number'; area = 'ADDRESS'; name='FAX_NUMBER'; options = @('default','create','update') }
        @{ displayName = 'first_name'; area = 'ADDRESS'; name='FIRSTNAME'; options = @('default','create','update') }
        @{ displayName = 'floor'; area = 'ADDRESS'; name='FLOOR'; options = @() }
        @{ displayName = 'FLOOR_P'; area = 'ADDRESS'; name='FLOOR_P'; options = @('create','update') }
        @{ displayName = 'full_name'; area = 'ADDRESS'; name='FULLNAME'; options = @('default','create','update') }
        @{ displayName = 'FULLNAME_X'; area = 'ADDRESS'; name='FULLNAME_X'; options = @() }
        @{ displayName = 'function'; area = 'ADDRESS'; name='FUNCTION'; options = @('create','update') }
        @{ displayName = 'GLOBAL_LOCK_STATE'; area = 'ISLOCKED'; name='GLOBAL_LOCK_STATE'; options = @('default') }
        @{ displayName = 'vaild_to'; area = 'LOGONDATA'; name='GLTGB'; options = @('default','create','update') }
        @{ displayName = 'valid_from'; area = 'LOGONDATA'; name='GLTGV'; options = @('default','create','update') }
        @{ displayName = 'scn_permit_sap_gui_checkbox'; area = 'SNC'; name='GUIFLAG'; options = @('default','create','update') }
        @{ displayName = 'HOMECITYNO'; area = 'ADDRESS'; name='HOMECITYNO'; options = @() }
        @{ displayName = 'HOME_CITY'; area = 'ADDRESS'; name='HOME_CITY'; options = @() }
        @{ displayName = 'house'; area = 'ADDRESS'; name='HOUSE_NO'; options = @('default','create','update') }
        @{ displayName = 'HOUSE_NO2'; area = 'ADDRESS'; name='HOUSE_NO2'; options = @() }
        @{ displayName = 'HOUSE_NO3'; area = 'ADDRESS'; name='HOUSE_NO3'; options = @() }
        @{ displayName = 'INHOUSE_ML'; area = 'ADDRESS'; name='INHOUSE_ML'; options = @() }
        @{ displayName = 'initials'; area = 'ADDRESS'; name='INITIALS'; options = @('default','create','update') }
        @{ displayName = 'INITS_SIG'; area = 'ADDRESS'; name='INITS_SIG'; options = @() }
        @{ displayName = 'cost_center'; area = 'DEFAULTS'; name='KOSTL'; options = @('default','create','update') }
        @{ displayName = 'language'; area = 'ADDRESS'; name='LANGU'; options = @('default','create','update') }
        @{ displayName = 'LANGUCPISO'; area = 'ADDRESS'; name='LANGUCPISO'; options = @() }
        @{ displayName = 'LANGUP_ISO'; area = 'ADDRESS'; name='LANGUP_ISO'; options = @() }
        @{ displayName = 'LANGU_CR_P'; area = 'ADDRESS'; name='LANGU_CR_P'; options = @() }
        @{ displayName = 'LANGU_ISO'; area = 'ADDRESS'; name='LANGU_ISO'; options = @() }
        @{ displayName = 'LANGU_P'; area = 'ADDRESS'; name='LANGU_P'; options = @() }
        @{ displayName = 'last_name'; area = 'ADDRESS'; name='LASTNAME'; options = @('default','create','update') }
        @{ displayName = 'LIC_TYPE'; area = 'UCLASS'; name='LIC_TYPE'; options = @('default','create','update') }
        @{ displayName = 'LOCAL_LOCK_STATE'; area = 'ISLOCKED'; name='LOCAL_LOCK_STATE'; options = @('default') }
        @{ displayName = 'LOCATION'; area = 'ADDRESS'; name='LOCATION'; options = @() }
        @{ displayName = 'LOCK_STATE'; area = 'ISLOCKED'; name='LOCK_STATE'; options = @('default','lock','unlock') }
        @{ displayName = 'LTIME'; area = 'LOGONDATA'; name='LTIME'; options = @() }
        @{ displayName = 'middle_name'; area = 'ADDRESS'; name='MIDDLENAME'; options = @('default','create','update') }
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
        @{ displayName = 'postal_code'; area = 'ADDRESS'; name='POSTL_COD1'; options = @('default','create','update') }
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
        @{ displayName = 'ROOM_NO_P'; area = 'ADDRESS'; name='ROOM_NO_P'; options = @('create','update') }
        @{ displayName = 'SECONDNAME'; area = 'ADDRESS'; name='SECONDNAME'; options = @() }
        @{ displayName = 'SECURITY_POLICY'; area = 'LOGONDATA'; name='SECURITY_POLICY'; options = @('default') }
        @{ displayName = 'SORT1'; area = 'ADDRESS'; name='SORT1'; options = @() }
        @{ displayName = 'SORT1_P'; area = 'ADDRESS'; name='SORT1_P'; options = @() }
        @{ displayName = 'SORT2'; area = 'ADDRESS'; name='SORT2'; options = @() }
        @{ displayName = 'SORT2_P'; area = 'ADDRESS'; name='SORT2_P'; options = @() }
        @{ displayName = 'delete_after_output'; area = 'DEFAULTS'; name='SPDA'; options = @('default','create','update') }
        @{ displayName = 'output_immediately'; area = 'DEFAULTS'; name='SPDB'; options = @('default','create','update') }
        @{ displayName = 'spool_output_service'; area = 'DEFAULTS'; name='SPLD'; options = @('default','create','update') }
        @{ displayName = 'SPEC_VERS'; area = 'UCLASS'; name='SPEC_VERS'; options = @() }
        @{ displayName = 'SPLG'; area = 'DEFAULTS'; name='SPLG'; options = @() }
        @{ displayName = 'START_MENU'; area = 'DEFAULTS'; name='START_MENU'; options = @() }
        @{ displayName = 'STCOD'; area = 'DEFAULTS'; name='STCOD'; options = @() }
        @{ displayName = 'street'; area = 'ADDRESS'; name='STREET'; options = @('default','create','update') }
        @{ displayName = 'STREET_NO'; area = 'ADDRESS'; name='STREET_NO'; options = @() }
        @{ displayName = 'STR_ABBR'; area = 'ADDRESS'; name='STR_ABBR'; options = @() }
        @{ displayName = 'STR_SUPPL1'; area = 'ADDRESS'; name='STR_SUPPL1'; options = @() }
        @{ displayName = 'STR_SUPPL2'; area = 'ADDRESS'; name='STR_SUPPL2'; options = @() }
        @{ displayName = 'STR_SUPPL3'; area = 'ADDRESS'; name='STR_SUPPL3'; options = @() }
        @{ displayName = 'SUBSTITUTE_FROM'; area = 'UCLASS'; name='SUBSTITUTE_FROM'; options = @() }
        @{ displayName = 'SUBSTITUTE_UNTIL'; area = 'UCLASS'; name='SUBSTITUTE_UNTIL'; options = @() }
        @{ displayName = 'SYSID'; area = 'UCLASS'; name='SYSID'; options = @() }
        @{ displayName = 'TAXJURCODE'; area = 'ADDRESS'; name='TAXJURCODE'; options = @() }
        @{ displayName = 'telephone_extension'; area = 'ADDRESS'; name='TEL1_EXT'; options = @('default','create','update') }
        @{ displayName = 'telephone_number'; area = 'ADDRESS'; name='TEL1_NUMBR'; options = @('default','create','update') }
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
        @{ displayName = 'AGR_TEXT'; name='AGR_TEXT'; options = @('default') }
        @{ displayName = 'FROM_DATE'; name='FROM_DAT'; options = @('default') }
        @{ displayName = 'TO_DATE'; name='TO_DAT'; options = @('default') }
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
    ProfileHT = [System.Collections.ArrayList]@()
    Profile = @(
        @{ displayName = 'PROFN'; name='PROFN'; options = @('default','key') }
        @{ displayName = 'TYP'; name='TYP'; options = @('default') }
    )
    ParameterHT = [System.Collections.ArrayList]@()
    Parameter = @(
        @{ displayName = 'PARAMID'; name='PARAMID'; options = @('default','key') }
        @{ displayName = 'PARTEXT'; name='PARTEXT'; options = @('default') }
    )
}

$Global:Properties.User | ForEach-Object { $Global:Properties.UserHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.UserHT = $Global:Properties.UserHT | Group-Object name -AsHashTable
$Global:Properties.User | ForEach-Object { $Global:Properties.UserInvertHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.UserInvertHT = $Global:Properties.UserInvertHT | Group-Object displayName -AsHashTable
$Global:Properties.UserRole | ForEach-Object { $Global:Properties.UserRoleHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.UserRoleHT = $Global:Properties.UserRoleHT | Group-Object name -AsHashTable
$Global:Properties.UserParameter | ForEach-Object { $Global:Properties.UserParameterHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.UserParameterHT = $Global:Properties.UserParameterHT | Group-Object name -AsHashTable
$Global:Properties.UserProfile | ForEach-Object { $Global:Properties.UserProfileHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.UserProfileHT = $Global:Properties.UserProfileHT | Group-Object name -AsHashTable
$Global:Properties.Role | ForEach-Object { $Global:Properties.RoleHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.RoleHT = $Global:Properties.RoleHT | Group-Object name -AsHashTable
$Global:Properties.Profile | ForEach-Object { $Global:Properties.ProfileHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.ProfileHT = $Global:Properties.ProfileHT | Group-Object name -AsHashTable
$Global:Properties.Parameter | ForEach-Object { $Global:Properties.ParameterHT.Add([PSCustomObject]$_) > $null }
$Global:Properties.ParameterHT = $Global:Properties.ParameterHT | Group-Object name -AsHashTable

$Global:User = [System.Collections.ArrayList]@()
$Global:UserRole = [System.Collections.ArrayList]@()
$Global:UserParameter = [System.Collections.ArrayList]@()
$Global:UserProfile = [System.Collections.ArrayList]@()

function Idm-UsersRead {
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
                $i = $returnData.Count
                Log info "Retrieving User Details"

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
                        try { $obj[$Global:Properties.UserHT[$prop.Metadata.Name].displayName] = $export.GetValue($prop.Metadata.Name) } catch { Log warn "ADDRESS UCLASS property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
                    }

                    #SNC (Export)
                    $export = $bapiFunctionCall2.GetObject('SNC')
                    foreach($prop in $export) {
                        try { $obj[$Global:Properties.UserHT[$prop.Metadata.Name].displayName] = $export.GetValue($prop.Metadata.Name) } catch { Log warn "SNC property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
                    }

                    #Logondata (Export)
                    $export = $bapiFunctionCall2.GetObject('LOGONDATA')
                    foreach($prop in $export) {
                        try { $obj[$Global:Properties.UserHT[$prop.Metadata.Name].displayName] = $export.GetValue($prop.Metadata.Name) } catch { Log warn "LOGONDATA property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
                    }

                    #Defaults (Export)
                    $export = $bapiFunctionCall2.GetObject('DEFAULTS')
                    foreach($prop in $export) {
                        try { $obj[$Global:Properties.UserHT[$prop.Metadata.Name].displayName] = $export.GetValue($prop.Metadata.Name) } catch { Log warn "DEFAULTS property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
                    }
                    
                    #UCLASS (Export)
                    $export = $bapiFunctionCall2.GetObject('UCLASS')
                    foreach($prop in $export) {
                        try { $obj[$Global:Properties.UserHT[$prop.Metadata.Name].displayName] = $export.GetValue($prop.Metadata.Name) } catch { Log warn "UCLASS property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
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
                            try { $table_obj[$Global:Properties.UserRoleHT[$prop.Metadata.Name].displayName] = $row.GetValue($prop.Metadata.Name) } catch { Log warn "User Roles property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
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
                            try { $table_obj[$Global:Properties.UserParameterHT[$prop.Metadata.Name].displayName] = $row.GetValue($prop.Metadata.Name) } catch { Log warn "User Parameters property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
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
                            try { $table_obj[$Global:Properties.UserProfileHT[$prop.Metadata.Name].displayName] = $row.GetValue($prop.Metadata.Name) } catch { Log warn "User Profiles property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
                        }
                        [void]$Global:UserProfile.Add([PSCustomObject]$table_obj);
                    } 
                    
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

function Idm-UserCreate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "User"

    if ($GetMeta) {
        #
        # Get meta data
        #
        @{
            semantics = 'create'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayname; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('create') } | ForEach-Object {
                    @{ name = $_.displayName; allowance = 'prohibited' }
                }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params = ConvertFrom-Json2 $SystemParams
        $function_params   = ConvertFrom-Json2 $FunctionParams

        # Setup Connection
        $Global:Connection = Open-SAPConnection -SystemParams $system_params -FunctionParams $function_params
        
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayName
        
        $properties = $function_params.Clone()

        $function = 'BAPI_USER_CREATE1'
        LogIO info $function -In @system_params -Properties $properties
        $repository = $Global:Connection.Repository
        [SAP.Middleware.Connector.IRfcFunction]$userCreate = $repository.CreateFunction($function)
        
        [SAP.Middleware.Connector.IRfcStructure]$userPassword = $userCreate.GetStructure("PASSWORD")
        [SAP.Middleware.Connector.IRfcStructure]$userAddress = $userCreate.GetStructure("ADDRESS")
        [SAP.Middleware.Connector.IRfcStructure]$userDefaults = $userCreate.GetStructure("DEFAULTS")
        [SAP.Middleware.Connector.IRfcStructure]$userLogonData = $userCreate.GetStructure("LOGONDATA")
        [SAP.Middleware.Connector.IRfcStructure]$userSNC = $userCreate.GetStructure("SNC")
        [SAP.Middleware.Connector.IRfcStructure]$userUClass = $userCreate.GetStructure("UCLASS")

        foreach($prop in ([PSCustomObject]$properties).PSObject.properties) {
            try { $field = $Global:Properties.UserInvertHT[$prop.Name] } catch { throw "[$($prop.Name)] does not have a connector mapping for user create, skipping"}

            if($prop.Name -eq $key) {
                $userCreate.SetValue($field.name,$prop.Value)
                continue
            }
            
            #Check for Password Fields
            if($field.name -in @('BAPIPWD','CODVN')) {

            } else {
                #Parse by Area
                switch($field.area) {
                    #Parse Outside of switch     
                    {$_ -in @('BAPIPWD','CODVN')} { continue }
                    
                    #Datetime Fields
                    # Maybe needed for GLTGV and GLTGB

                    #Boolean Fields
                    {$_ -eq 'SNC' -and $field.name -eq 'GUIFLAG'} { $userSNC.SetValue('GUIFLAG',$prop.Value)}
                    
                    #Direct Mapping
                    'ADDRESS'   { $userAddress.SetValue($field.name,$prop.Value) }
                    'SNC'       { $userSNC.SetValue($field.name,$prop.Value) }
                    'DEFAULTS'  { $userDefaults.SetValue($field.name,$prop.Value) }
                    'LOGONDATA' { $userLogonData.SetValue($field.name,$prop.Value) }
                    'UCLASS'    { $userUClass.SetValue($field.name,$prop.Value) }
                    
                    
                    #Failback
                    default { LogIO warn $function -Out "[$($field.name)] in [$($field.area)] does not have a connector mapping for user create, skipping" }
                }
            }
        }

        #Parse password Options
        if($user.disable_password) {
            $userLogonData.SetValue('CODVN','X')
            $userPassword.SetValue('BAPIPWD','')
        } else {
            $userPassword.SetValue('BAPIPWD',$Properties.BAPIPWD)
        }

        $userCreate.Invoke($Global:Connection)
        Get-ReturnLog -Call $userCreate -Context $function
        LogIO info $function
    }

    Log info "Done"
}

function Idm-UserUpdate {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "User"

    if ($GetMeta) {
        #
        # Get meta data
        #
        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayname; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('create') } | ForEach-Object {
                    @{ name = $_.displayName; allowance = 'prohibited' }
                }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params = ConvertFrom-Json2 $SystemParams
        $function_params   = ConvertFrom-Json2 $FunctionParams

        # Setup Connection
        $Global:Connection = Open-SAPConnection -SystemParams $system_params -FunctionParams $function_params
        
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayName
        
        $properties = $function_params.Clone()

        $function = 'BAPI_USER_CHANGE'
        LogIO info $function -In @system_params -Properties $properties
        $repository = $Global:Connection.Repository
        [SAP.Middleware.Connector.IRfcFunction]$userUpdate = $repository.CreateFunction($function)
        
        [SAP.Middleware.Connector.IRfcStructure]$userPassword = $userUpdate.GetStructure("PASSWORD")
        [SAP.Middleware.Connector.IRfcStructure]$userAddress = $userUpdate.GetStructure("ADDRESS")
        [SAP.Middleware.Connector.IRfcStructure]$userDefaults = $userUpdate.GetStructure("DEFAULTS")
        [SAP.Middleware.Connector.IRfcStructure]$userLogonData = $userUpdate.GetStructure("LOGONDATA")
        [SAP.Middleware.Connector.IRfcStructure]$userSNC = $userUpdate.GetStructure("SNC")
        [SAP.Middleware.Connector.IRfcStructure]$userUClass = $userUpdate.GetStructure("UCLASS")

        foreach($prop in ([PSCustomObject]$properties).PSObject.properties) {
            try { $field = $Global:Properties.UserInvertHT[$prop.Name] } catch { throw "[$($prop.Name)] does not have a connector mapping for user create, skipping"}

            if($prop.Name -eq $key) {
                $userUpdate.SetValue($field.name,$prop.Value)
                continue
            }
            
            #Check for Password Fields
            if($field.name -in @('BAPIPWD','CODVN')) {

            } else {
                #Parse by Area
                switch($field.area) {
                    #Parse Outside of switch     
                    {$_ -in @('BAPIPWD','CODVN')} { continue }
                    
                    #Datetime Fields
                    # Maybe needed for GLTGV and GLTGB

                    #Boolean Fields
                    {$_ -eq 'SNC' -and $field.name -eq 'GUIFLAG'} { $userSNC.SetValue('GUIFLAG',($PropValue -eq 'X'))}
                    
                    #Direct Mapping
                    'ADDRESS'   { $userAddress.SetValue($field.name,$prop.Value); Log info "[$($field.Name)] - [$($prop.Value)]" }
                    'SNC'       { $userSNC.SetValue($field.name,$prop.Value) }
                    'DEFAULTS'  { $userDefaults.SetValue($field.name,$prop.Value) }
                    'LOGONDATA' { $userLogonData.SetValue($field.name,$prop.Value) }
                    'UCLASS'    { $userUClass.SetValue($field.name,$prop.Value) }
                    
                    #Failback
                    default { LogIO warn $function -Out "[$($field.name)] in [$($field.area)] does not have a connector mapping for user create, skipping" }
                }
            }
        }

        #Parse password Options
        if($user.disable_password) {
            $userLogonData.SetValue('CODVN','X')
            $userPassword.SetValue('BAPIPWD','')
        } elseif ($user.BAPIPWD.length -gt 0) {
            $userPassword.SetValue('BAPIPWD',$Properties.BAPIPWD)
        }
        
        $userUpdate.Invoke($Global:Connection)
        Get-ReturnLog -Call $userUpdate -Context $function
        LogIO info $function
    }

    Log info "Done"
}

function Idm-UserLock {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "User"

    if ($GetMeta) {
        #
        # Get meta data
        #
        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayname; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('lock') } | ForEach-Object {
                    @{ name = $_.displayName; allowance = 'prohibited' }
                }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params = ConvertFrom-Json2 $SystemParams
        $function_params   = ConvertFrom-Json2 $FunctionParams

        # Setup Connection
        $Global:Connection = Open-SAPConnection -SystemParams $system_params -FunctionParams $function_params
        
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayName
        
        $properties = $function_params.Clone()

        $function = 'BAPI_USER_LOCK'
        LogIO info $function -In @system_params -Properties $properties
        $repository = $Global:Connection.Repository
        [SAP.Middleware.Connector.IRfcFunction]$userUpdate = $repository.CreateFunction($function)

        foreach($prop in ([PSCustomObject]$properties).PSObject.properties) {
            try { $field = $Global:Properties.UserInvertHT[$prop.Name] } catch { throw "[$($prop.Name)] does not have a connector mapping for user create, skipping"}

            if($prop.Name -eq $key) {
                $userUpdate.SetValue($field.name,$prop.Value)
                continue
            }
        }

        $userUpdate.Invoke($Global:Connection)
        Get-ReturnLog -Call $userUpdate -Context $function
        LogIO info $function
    }

    Log info "Done"
}

function Idm-UserUnlock {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "User"

    if ($GetMeta) {
        #
        # Get meta data
        #
        @{
            semantics = 'update'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayname; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('lock') } | ForEach-Object {
                    @{ name = $_.displayName; allowance = 'prohibited' }
                }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params = ConvertFrom-Json2 $SystemParams
        $function_params   = ConvertFrom-Json2 $FunctionParams

        # Setup Connection
        $Global:Connection = Open-SAPConnection -SystemParams $system_params -FunctionParams $function_params
        
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayName
        
        $properties = $function_params.Clone()

        $function = 'BAPI_USER_UNLOCK'
        LogIO info $function -In @system_params -Properties $properties
        $repository = $Global:Connection.Repository
        [SAP.Middleware.Connector.IRfcFunction]$userUpdate = $repository.CreateFunction($function)

        foreach($prop in ([PSCustomObject]$properties).PSObject.properties) {
            try { $field = $Global:Properties.UserInvertHT[$prop.Name] } catch { throw "[$($prop.Name)] does not have a connector mapping for user create, skipping"}

            if($prop.Name -eq $key) {
                $userUpdate.SetValue($field.name,$prop.Value)
                continue
            }
        }

        $userUpdate.Invoke($Global:Connection)
        Get-ReturnLog -Call $userUpdate -Context $function
        LogIO info $function
    }

    Log info "Done"
}

function Idm-UserDelete {
    param (
        # Operations
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "User"

    if ($GetMeta) {
        #
        # Get meta data
        #
        @{
            semantics = 'delete'
            parameters = @(
                @{ name = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayname; allowance = 'mandatory' }

                $Global:Properties.$Class | Where-Object { !$_.options.Contains('key') -and !$_.options.Contains('delete') } | ForEach-Object {
                    @{ name = $_.displayName; allowance = 'prohibited' }
                }
            )
        }
    }
    else {
        #
        # Execute function
        #

        $system_params = ConvertFrom-Json2 $SystemParams
        $function_params   = ConvertFrom-Json2 $FunctionParams

        # Setup Connection
        $Global:Connection = Open-SAPConnection -SystemParams $system_params -FunctionParams $function_params
        
        $key = ($Global:Properties.$Class | Where-Object { $_.options.Contains('key') }).displayName
        $properties = $function_params.Clone()
        $function = 'BAPI_USER_DELETE'
        LogIO info $function -In @system_params -Properties $properties
        $repository = $Global:Connection.Repository
        [SAP.Middleware.Connector.IRfcFunction]$userDelete = $repository.CreateFunction($function)
        $userDelete.SetValue('USERNAME',$Properties.$key)
        $userDelete.Invoke($Global:Connection)
        Get-ReturnLog -Call $userDelete -Context $function

        LogIO info $function
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
            Idm-UsersRead -SystemParams $SystemParams | Out-Null
        } 

        foreach($item in $Global:UserRole) {            
            $item | Select-Object $displayProperties
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
            Idm-UsersRead -SystemParams $SystemParams | Out-Null
        } 

        foreach($item in $Global:UserParameter) {            
            $item | Select-Object $displayProperties
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
            Idm-UsersRead -SystemParams $SystemParams | Out-Null
        } 

        foreach($item in $Global:UserProfile) {            
            $item | Select-Object $displayProperties
        }
    }

    Log info "Done"
}

function Idm-RolesRead {
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
            ([PSCustomObject]$table_obj) | Select-Object $displayProperties
        } 
    }

    Log info "Done"
}

function Idm-ProfilesRead {
    param (
        [switch] $GetMeta,
        [string] $SystemParams,
        [string] $FunctionParams
    )
    
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "Profile"
    
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

        $displayProperties = $properties | ForEach-Object { $Global:Properties.ProfileHT[$_].displayName }
        
        Log info "Retrieving Profile List"
        $ret = Read-Table -destination $Global:Connection -QueryTable "USR10" -Fields $properties
        
        foreach($item in $ret) {
            $obj = @{}
            foreach($prop in $item.PSObject.Properties) {
                try { $obj[($Global:Properties.ProfileHT[$prop.Name]).displayName] = $prop.Value } catch { Log warn "Profiles property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
            }
            ([PSCustomObject]$obj) | Select-Object $displayProperties
        }
    }

    Log info "Done"
}

function Idm-ParametersRead {
    param (
        [switch] $GetMeta,
        [string] $SystemParams,
        [string] $FunctionParams
    )
    
    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $Class = "Parameter"
    
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

        $displayProperties = $properties | ForEach-Object { $Global:Properties.ParameterHT[$_].displayName }
        
        Log info "Retrieving Parameter List"
        $ret = Read-Table -destination $Global:Connection -QueryTable "TPARA" -Fields $properties
        
        foreach($item in $ret) {
            $obj = @{}
            foreach($prop in $item.PSObject.Properties) {
                try { $obj[($Global:Properties.ParameterHT[$prop.Name]).displayName] = $prop.Value  } catch { Log warn "Parameters property [$($prop.Metadata.Name)] not defined in configuration, skipping" }
            }
            ([PSCustomObject]$obj) | Select-Object $displayProperties
        }
    }

    Log info "Done"
}

#
# Supplemental functions
#

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

function Read-Table {
    param (
        [parameter(Mandatory = $true)]$destination,
        [parameter(Mandatory = $true)]$QueryTable,
        [parameter(Mandatory = $false)]$Delimiter = ";",
        [parameter(Mandatory = $false)][System.Collections.ArrayList]$Options,
        [parameter(Mandatory = $false)][String[]]$Fields
    )
    
    try {
        $repository = $destination.Repository
        [SAP.Middleware.Connector.IRfcFunction]$bapiReadRfcTable = $repository.CreateFunction("RFC_READ_TABLE")
        [SAP.Middleware.Connector.IRfcTable]$SAPOptions = $bapiReadRfcTable.GetTable("OPTIONS")
        [SAP.Middleware.Connector.IRfcTable]$SAPFields = $bapiReadRfcTable.GetTable("FIELDS")
        $bapiReadRfcTable.SetValue("QUERY_TABLE", $QueryTable)
        $bapiReadRfcTable.SetValue("DELIMITER", $Delimiter)

        if ($null -ne $Fields) {
            foreach ($line in $Fields) {
                $SAPFields.Append()
                $SAPFields.SetValue("FIELDNAME", $line) 
            }
        }     

        if ($null -ne $Options) {
            foreach ($line in $Options) {
                $SAPOptions.Append()
                $SAPOptions.SetValue("TEXT", $line.TEXT) 
            }
        }     
        
        $bapiReadRfcTable.Invoke($destination)

        $sap_fields = New-Object System.Collections.ArrayList
        [SAP.Middleware.Connector.IRfcTable]$FIELDS = $bapiReadRfcTable.GetTable("FIELDS")
        foreach ($record in $FIELDS) {
            $row = [PSCustomObject]@{
                "FIELDNAME"  = $record.GetValue("FIELDNAME")
                "TYPE"  = $record.GetValue("TYPE")
            }
            $sap_fields.Add($row) > $null
        }
        $header = "";
        foreach ($line in $sap_fields) {
            $header = $header + $line.FIELDNAME + ";"
        }
        $header = $header.TrimEnd(";");

        $sap_table = New-Object System.Collections.ArrayList
        [SAP.Middleware.Connector.IRfcTable]$DATA = $bapiReadRfcTable.GetTable("DATA")
        $firstrow = $true
        foreach ($record in $DATA) {
            $row = [PSCustomObject]@{
                "WA"  = $record.GetValue("WA")
            }
            $sap_table.Add($row) > $null
        }
        $output = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($line in $sap_table.wa) {
            if ($firstrow -eq $true)
            {
                $sap_table.Add($row) > $null
                $firstrow = $false
                $output.Add([PSCustomObject]@{
                    column  = $header
                })
            }
            $array = $line.Split(";") 
            $line = $array.trim() -join ";"
            $output.Add([PSCustomObject]@{
                column  = $line
            })
        }
        $sap_table = ConvertFrom-Csv -Delimiter ";" -InputObject $output.column

        foreach ($row in $sap_table) {
            foreach ($sapfield in $sap_fields) {
                $val = $row."$($sapfield.FIELDNAME)"
                if ($sapfield.TYPE -eq "D")
                {
                    if ($val -eq "00000000") {
                        $val = $null
                    }
                    else {
                        $val = [datetime]::ParseExact($val, 'yyyyMMdd', $null)
                    }
                }
                $row."$($sapfield.FIELDNAME)" = $val
            }
        }
    }
    catch {
        $msg = "Could not get the content of $($QueryTable), message: $($_.Exception.Message)"
        Log error $msg
        throw $msg
    }
    finally {
        Write-Output $sap_table
    }

}
function Get-ReturnLog {
    param (
        $Call,
        $Context
    )
    #Get Return Table
    $table = $Call.GetTable("RETURN")  

    foreach($row in $table) {
        $rowObj = @{}
        foreach($prop in $row) {
            $rowObj[$prop.Metadata.Name] = $row.GetValue($prop.Metadata.Name)
        }

        $type = "info"
        if($rowObj.Type -eq 'E') { 
            $type = "error"
            Log error $rowObj.MESSAGE
        }

        LogIO $type $Context -Out $rowObj
    } 
}
    
#
# Configuration
#

$configScenarios = @'
[{"name":"Default","description":"Default Configuration","version":"1.0","createTime":1737388671671,
"modifyTime":1737388671671,"name_values":[{"name":"Client","value":null},{"name":"Hostname","value":
null},{"name":"Password","value":null},{"name":"SAPdllDirectoryPath","value":null},{"name":"SysId","
value":null},{"name":"SysNr","value":null},{"name":"Username","value":null},{"name":"collections","v
alue":["Parameters","Profiles","Roles","UserParameters","UserProfiles","UserRoles","Users"]},{"name"
:"nr_of_sessions","value":null},{"name":"sessions_idle_timeout","value":null}],"collections":[{"col_
name":"Users","fields":[{"field_name":"Username","field_type":"string","include":true,"field_format"
:"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fiel
d_name":"account_id","field_type":"string","include":true,"field_format":"","field_source":"data","j
avascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"building","field_ty
pe":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"re
ference":false,"ref_col_fields":[]},{"field_name":"city","field_type":"string","include":true,"field
_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]
},{"field_name":"department","field_type":"string","include":true,"field_format":"","field_source":"
data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"E_MAIL","fi
eld_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":
[],"reference":false,"ref_col_fields":[]},{"field_name":"facsimile_extension","field_type":"string",
"include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fals
e,"ref_col_fields":[]},{"field_name":"facsimile_number","field_type":"string","include":true,"field_
format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]}
,{"field_name":"first_name","field_type":"string","include":true,"field_format":"","field_source":"d
ata","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"full_name","
field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col
":[],"reference":false,"ref_col_fields":[]},{"field_name":"GLOBAL_LOCK_STATE","field_type":"string",
"include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fals
e,"ref_col_fields":[]},{"field_name":"vaild_to","field_type":"string","include":true,"field_format":
"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field
_name":"valid_from","field_type":"string","include":true,"field_format":"","field_source":"data","ja
vascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"scn_permit_sap_gui_c
heckbox","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"
","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"house","field_type":"string","i
nclude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,
"ref_col_fields":[]},{"field_name":"initials","field_type":"string","include":true,"field_format":""
,"field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_n
ame":"cost_center","field_type":"string","include":true,"field_format":"","field_source":"data","jav
ascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"language","field_type
":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"refe
rence":false,"ref_col_fields":[]},{"field_name":"last_name","field_type":"string","include":true,"fi
eld_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields"
:[]},{"field_name":"LIC_TYPE","field_type":"string","include":true,"field_format":"","field_source":
"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"LOCAL_LOCK
_STATE","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":""
,"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"LOCK_STATE","field_type":"string
","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":fa
lse,"ref_col_fields":[]},{"field_name":"middle_name","field_type":"string","include":true,"field_for
mat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"
field_name":"NO_PASSWORD_LOCK_STATE","field_type":"string","include":true,"field_format":"","field_s
ource":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"sec
ure_network_communication","field_type":"string","include":true,"field_format":"","field_source":"da
ta","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"postal_code",
"field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_co
l":[],"reference":false,"ref_col_fields":[]},{"field_name":"room_number","field_type":"string","incl
ude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"re
f_col_fields":[]},{"field_name":"SECURITY_POLICY","field_type":"string","include":true,"field_format
":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fie
ld_name":"delete_after_output","field_type":"string","include":true,"field_format":"","field_source"
:"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"output_im
mediately","field_type":"string","include":true,"field_format":"","field_source":"data","javascript"
:"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"spool_output_service","field_
type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"
reference":false,"ref_col_fields":[]},{"field_name":"street","field_type":"string","include":true,"f
ield_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields
":[]},{"field_name":"telephone_extension","field_type":"string","include":true,"field_format":"","fi
eld_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name"
:"telephone_number","field_type":"string","include":true,"field_format":"","field_source":"data","ja
vascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"WRONG_LOGON_LOCK_STA
TE","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","re
f_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"CLIENT","field_type":"string","inclu
de":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref
_col_fields":[]}],"key":"Username","display":"Username","name_values":[],"sys_nn":[],"source":"data"
},{"col_name":"UserParameters","fields":[{"field_name":"Username","field_type":"string","include":tr
ue,"field_format":"","field_source":"data","javascript":"","ref_col":["Users"],"reference":false,"re
f_col_fields":[]},{"field_name":"PARID","field_type":"string","include":true,"field_format":"","fiel
d_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"
PARVA","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"",
"ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"PARTXT","field_type":"string","in
clude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"
ref_col_fields":[]},{"field_name":"parid","field_type":"string","include":true,"field_format":"","fi
eld_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name"
:"parva","field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"
","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"partxt","field_type":"string","
include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false
,"ref_col_fields":[]}],"key":"","display":"","name_values":[],"sys_nn":[],"source":"data"},{"col_nam
e":"UserProfiles","fields":[{"field_name":"Username","field_type":"string","include":true,"field_for
mat":"","field_source":"data","javascript":"","ref_col":["Users"],"reference":false,"ref_col_fields"
:[]},{"field_name":"BAPIPROF","field_type":"string","include":true,"field_format":"","field_source":
"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"BAPIPTEXT"
,"field_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_c
ol":[],"reference":false,"ref_col_fields":[]},{"field_name":"BAPIAKTPS","field_type":"string","inclu
de":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref
_col_fields":[]},{"field_name":"BAPITYPE","field_type":"string","include":true,"field_format":"","fi
eld_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name"
:"bapiprof","field_type":"string","include":true,"field_format":"","field_source":"data","javascript
":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"bapiptext","field_type":"str
ing","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference"
:false,"ref_col_fields":[]},{"field_name":"bapiaktps","field_type":"string","include":true,"field_fo
rmat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{
"field_name":"bapitype","field_type":"string","include":true,"field_format":"","field_source":"data"
,"javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]}],"key":"","display":"","name_va
lues":[],"sys_nn":[],"source":"data"},{"col_name":"UserRoles","fields":[{"field_name":"Username","fi
eld_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":
["Users"],"reference":false,"ref_col_fields":[]},{"field_name":"AGR_NAME","field_type":"string","inc
lude":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"r
ef_col_fields":[]},{"field_name":"FRM_DATE","field_type":"string","include":true,"field_format":"","
field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_nam
e":"TO_DATE","field_type":"string","include":true,"field_format":"","field_source":"data","javascrip
t":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"field_name":"ORG_FLAG","field_type":"str
ing","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference"
:false,"ref_col_fields":[]},{"field_name":"FROM_DATE","field_type":"string","include":true,"field_fo
rmat":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{
"field_name":"AGR_TEXT","field_type":"string","include":true,"field_format":"","field_source":"data"
,"javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]}],"key":"","display":"","name_va
lues":[],"sys_nn":[],"source":"data"},{"col_name":"Parameters","fields":[{"field_name":"PARAMID","fi
eld_type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":
[],"reference":false,"ref_col_fields":[]},{"field_name":"PARTEXT","field_type":"string","include":tr
ue,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_f
ields":[]}],"key":"PARAMID","display":"PARAMID","name_values":[],"sys_nn":[],"source":"data"},{"col_
name":"Profiles","fields":[{"field_name":"PROFN","field_type":"string","include":true,"field_format"
:"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]},{"fiel
d_name":"TYP","field_type":"string","include":true,"field_format":"","field_source":"data","javascri
pt":"","ref_col":[],"reference":false,"ref_col_fields":[]}],"key":"PROFN","display":"PROFN","name_va
lues":[],"sys_nn":[],"source":"data"},{"col_name":"Roles","fields":[{"field_name":"AGR_NAME","field_
type":"string","include":true,"field_format":"","field_source":"data","javascript":"","ref_col":[],"
reference":false,"ref_col_fields":[]},{"field_name":"FLAG_COLL","field_type":"string","include":true
,"field_format":"","field_source":"data","javascript":"","ref_col":[],"reference":false,"ref_col_fie
lds":[]},{"field_name":"TEXT","field_type":"string","include":true,"field_format":"","field_source":
"data","javascript":"","ref_col":[],"reference":false,"ref_col_fields":[]}],"key":"AGR_NAME","displa
y":"TEXT","name_values":[],"sys_nn":[],"source":"data"}]}]
'@
