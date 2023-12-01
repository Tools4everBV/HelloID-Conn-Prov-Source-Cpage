# Getting variables from configuration tab
$config = ConvertFrom-Json $configuration
Try{
    # Initializing Oracle connexion
    [System.Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient") | out-null
    $conn = [System.Data.Odbc.OdbcConnection]::new()
    $conn.connectionstring = $config.connectionstring
    $conn.open()

    # Querying Persons
    $sqlQuerySelectPersons = "SELECT
    AGE.NUAGAG AS ExternalId,
    AGE.NPATAG AS LastNameBirth,
    AGE.NUSUAG AS LastName,
    AGE.PRENAG AS FirstName,
    AGE.NUSUAG || ' ' || AGE.PRENAG || '-' || AGE.NUAGAG AS DisplayName,
    AGE.TITRAG AS Civility,
    PSPEB2.CODESSPEB2 AS B2_Code,
    PINFPS.NRPPSINFPS AS Rpps_Code
    FROM
        AGE,PSPEB2,PINFPS
    WHERE AGE.NUAGAG = PSPEB2.MATRISPEB2(+)
        AND AGE.NUAGAG = PINFPS.MATRIINFPS(+)
    "

    $cmdSelectPersons = [System.Data.Odbc.OdbcCommand]::new($sqlQuerySelectPersons, $conn)
    $da = [System.Data.Odbc.OdbcDataAdapter]::new($cmdSelectPersons)
    $dtPersons = [System.Data.datatable]::new()
    $null = $da.fill($dtPersons)
    $persons += $dtPersons
    $persons = $persons | Select-Object -Property * -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors

    # Querying ADELI Codes
    $sqlAdeli = "SELECT PER.PADELI.MATRIADELI,
    PER.PADELI.NUMERADELI,
    PER.PADELI.DATEDADELI
    FROM 
        PER.PADELI"
    $cmdAdeli = [System.Data.Odbc.OdbcCommand]::new($sqlAdeli, $conn)
    $da = [System.Data.Odbc.OdbcDataAdapter]::new($cmdAdeli)
    $dtAdeli = [System.Data.datatable]::new()
    $null = $da.fill($dtAdeli)
    $CodesAdeli = @()
    $CodesAdeli += $dtAdeli

    # Indexing list of Adeli codes numbers
    $CodesAdeli = $CodesAdeli | Select-Object -Property * -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors
    $CodesAdeliGrouped = $CodesAdeli  | Group-Object -Property "MATRIADELI" -AsHashTable -AsString

    # Querying Contracts
    $sqlQuerySelectContracts = "SELECT DISTINCT
    PAFREP.UFREPPAFREP AS Department_Code,
    PAFREP.AGENTPAFREP AS Employee_ID,
    TO_CHAR (PAFREP.DTDEBPAFREP,'mm-dd-yyyy') AS Start_Date,
    TO_CHAR (PAFREP.DTFINPAFREP,'mm-dd-yyyy') AS End_Date,
    PAFREP.IDENTPAFREP AS Contract_ID,
    PAMT2.COFAAMT||PAMT2.COSFAMT||PAMT2.CODEAMT AS Business_Code,
    PCME2.LIOFCME AS Job_Title,
    U.LIBPUF AS Department_NameLong
    FROM
        PAFREP,
        PAMT2,
        PCME2,
        UFP U
    WHERE (PAFREP.DTFINPAFREP > '01/06/2021' OR PAFREP.DTFINPAFREP is NULL)
	    AND PAFREP.AGENTPAFREP = PAMT2.AGENAMT
        AND (PAMT2.DTFIAMT IS NULL OR TO_CHAR(PAMT2.DTFIAMT,'yyyymmdd') >= TO_CHAR(sysdate,'yyyymmdd'))
        AND (PAMT2.CODEAMT=PCME2.CODECME(+) and PAMT2.COFAAMT=PCME2.COFACME(+) and PAMT2.COSFAMT=PCME2.COSFCME(+))
        AND PAFREP.UFREPPAFREP = U.NUFPUF
        AND PAMT2.MEPRAMT = 'O'
    "

    $cmdSelectContracts = [System.Data.Odbc.OdbcCommand]::new($sqlQuerySelectContracts, $conn)
    $da = [System.Data.Odbc.OdbcDataAdapter]::new($cmdSelectContracts)
    $dtContracts = [System.Data.datatable]::new()
    $null = $da.fill($dtContracts)
    $contracts = @()
    $contracts += $dtContracts
    $contracts = $contracts | Select-Object -Property * -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors
    $conn.close()
    
    # Creating the person object                         
    foreach($p in $persons){
        $person = @{
            ExternalId      = $p.externalId
            DisplayName     = $p.displayName
            FirstName       = $p.firstName
	    LastName	    = $p.lastName
            LastNameBirth   = $p.lastNameBirth
            AdelNumber      = $codesAdeliGrouped["$($p.ExternalId)"].NUMERADELI
            RPPSCode        = $p.Rpps_Code
            B2Code          = $p.B2_Code
            Source          = "CPAGE"
            Contracts       = [System.Collections.ArrayList]@()
        }
	# Adding contracts to person           
        foreach($c in $contracts){

            if($c.Employee_ID -eq $p.ExternalId){
                $contract = @{
                    ID                 = $c.Contract_ID
                    BusinessCode       = $c.Business_Code
                    JobTitle           = $c.Job_Title
                    DepartmentCode     = $c.Department_Code
                    DepartmentNameLong = $c.Department_NameLong
                    NumPeriode         = $c.Sequence_Number
                    StartDate          = $c.Start_Date
                    EndDate            = $c.End_Date
                    FTE                = $c.FTE
                }
                [void]$person.Contracts.Add($contract);
            }
        }
        If ($person.contracts){
            Write-Output ($person | ConvertTo-Json -Depth 50)
        }
    }
    Write-Information "Persons data imported"

}Catch{
    Write-error "Error when importing persons - $($_.Exception.Message)"
}
