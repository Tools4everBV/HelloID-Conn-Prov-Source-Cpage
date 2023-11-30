$config = ConvertFrom-Json $configuration
try{
    [System.Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")| out-null
    $conn = New-Object System.Data.Odbc.OdbcConnection
    $conn.connectionstring = $config.dsn
    $conn.open()

    ### Unités Fonctionnelles
    $sql = "SELECT DISTINCT * FROM UFP"
    $cmd = New-Object system.Data.Odbc.OdbcCommand($sql,$conn)
    $da = New-Object system.Data.Odbc.OdbcDataAdapter($cmd)
    $dt = New-Object system.Data.datatable
    $null = $da.fill($dt)
    $Departments += $dt
    $Departments = $Departments | Select-Object -Property * -ExcludeProperty RowError,RowState,Table,ItemArray,HasErrors 
    $conn.close()
    $result = @()
    
    foreach($item in $Departments){
        $department = @{
            ExternalId=$item.NUFPUF
            DisplayName=$item.LBRPUF
            Name=$item.LBRPUF 
        }
        Write-Output ($department | ConvertTo-Json -Depth 50)
        $result += $department 
    }
Write-Verbose -Verbose "$($result.count) departements imported"
}catch{
    Write-verbose -verbose "Error during UF importation - $error[0].Exception.Message"
}
