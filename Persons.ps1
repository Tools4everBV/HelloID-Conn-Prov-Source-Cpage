$config = ConvertFrom-Json $configuration
Try{
    [System.Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient") | out-null
    $conn = New-Object System.Data.Odbc.OdbcConnection
    $conn.connectionstring = $config.dsn
    $conn.open()

    ### Persons
    $sql = "SELECT
    AGE.NUAGAG AS ExternalId,
    AGE.NPATAG AS LastNameBirth,
    AGE.NUSUAG AS LastName,
    AGE.PRENAG AS FirstName,
    AGE.NUSUAG || ' ' || AGE.PRENAG || '-' || AGE.NUAGAG  AS DisplayName,
    AGE.TITRAG AS Civility,
    PSPEB2.CODESSPEB2 AS B2_Code,
    PINFPS.NRPPSINFPS AS Rpps_Code
    FROM
    AGE,PSPEB2,PINFPS
    WHERE AGE.NUAGAG = PSPEB2.MATRISPEB2(+)
    AND AGE.NUAGAG = PINFPS.MATRIINFPS(+)
    "
   
    $cmd = New-Object system.Data.Odbc.OdbcCommand($sql,$conn)
    $da = New-Object system.Data.Odbc.OdbcDataAdapter($cmd)
    $dt = New-Object system.Data.datatable
    $null = $da.fill($dt)
    $persons += $dt
    $persons = $persons | Select-Object -Property * -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors 

    ### ADELI Codes
    $sql = "SELECT PER.PADELI.MATRIADELI,
    PER.PADELI.NUMERADELI,
    PER.PADELI.DATEDADELI
    FROM PER.PADELI"
    
    $cmd = New-Object system.Data.Odbc.OdbcCommand($sql,$conn)
    $da = New-Object system.Data.Odbc.OdbcDataAdapter($cmd)
    $dt = New-Object system.Data.datatable
    $null = $da.fill($dt)
    $CodesAdeli = @()
    $CodesAdeli += $dt
    $CodesAdeli = $CodesAdeli | Select-Object -Property * -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors | Group-Object MATRIADELI | Foreach-Object {$_.Group | Sort-Object DATEDADELI | Select-Object -Last 1} 

    ### Contracts
    $sql = "SELECT DISTINCT
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
    
    $cmd = New-Object system.Data.Odbc.OdbcCommand($sql,$conn)
    $da = New-Object system.Data.Odbc.OdbcDataAdapter($cmd)
    $dt = New-Object system.Data.datatable
    $null = $da.fill($dt)
    $Contracts = @()
    $Contracts += $dt
    $Contracts = $Contracts | Select-Object -Property * -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors 
    
    $conn.close()
    
    foreach($p in $persons){
        $person = @{};
        $person["ExternalId"] = $p.ExternalId
        $person["DisplayName"] = $p.DisplayName
        $person["FirstName"] = $p.FirstName
        $person["LastName"] = $p.LastName
        $person["LastNameBirth"] = $p.LastNameBirth
        $person["AdelNumber"] = ($CodesAdeli | where-object {$_.MATRIADELI -eq $p.ExternalId}).NUMERADELI
        $person["RPPSCode"] = $p.Rpps_Code
        $person["B2Code"] = $p.B2_Code
        $person["Source"] = "CPAGE"
        $person["Contracts"] = [System.Collections.ArrayList]@();
        foreach($c in $contracts){
            if($c.Employee_ID -eq $p.ExternalId){
                $contract = @{}; 
                $contract["ID"] = $c.Contract_ID
                $contract["BusinessCode"] = $c.Business_Code
                $contract["JobTitle"] = $c.Job_Title
                $contract["DepartmentCode"] = $c.Department_Code
                $contract["DepartmentNameLong"] = $c.Department_NameLong
                $contract["NumPeriode"] = $c.Sequence_Number
                $contract["StartDate"] = $c.Start_Date
                $contract["EndDate"] = $c.End_Date
                $contract["FTE"] = $c.FTE
                
                [void]$person.Contracts.Add($contract);
            }
        }
        If ($person.contracts){
            Write-Output ($person | ConvertTo-Json -Depth 50)
        }
    } 
    Write-Information "Persons data imported"

}Catch{
    Write-error "Erreur when importing persons - $error[0].Exception.Message"
}

