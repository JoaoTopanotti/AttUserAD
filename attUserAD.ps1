# Importa Módulo ActiveDirectory
Import-Module ActiveDirectory

$excel = Read-Host "Cole (botão direito) o caminho do .csv aqui"
$dadosExcel = Import-Csv -Path $excel -Delimiter ";" -Encoding UTF8

foreach ($usuario in $dadosExcel) {
    $SamAccountName = $usuario.Usuario
    $DisplayName = $usuario.NomeCompleto
    $Title = $usuario.Cargo
    $Office = $usuario.Unidade
    $Department = $usuario.Departamento
    $Description = $usuario.CentroCusto
    $employeeNumber = $usuario.Cpf
    $employeeID = $usuario.Matricula
    $matriculaGestor = $usuario.MatriculaGestor

    # Busca o usuário pelo CPF
    if (![string]::IsNullOrWhiteSpace($employeeNumber)) {
        $dadosUsuario = Get-ADUser -Filter "employeeNumber -eq '$employeeNumber'" -Properties *
    } else {
        Write-Warning "CPF ausente para o usuário $SamAccountName. Pulando..."
        continue
    }

    # Define o gestor do usuário
    $gestorDN = $null
    if (![string]::IsNullOrWhiteSpace($matriculaGestor)) {
        $gestorDN = Get-ADUser -Filter "employeeID -eq '$matriculaGestor'" -Properties * | Select-Object -ExpandProperty DistinguishedName
        if ($gestorDN) {
            Set-ADUser -Identity $dadosUsuario.SamAccountName -Manager $gestorDN
        }
    }

    # Atualiza o usuário no AD
    try {
        Set-ADUser -Identity $dadosUsuario.SamAccountName `
            -DisplayName $DisplayName `
            -Title $Title `
            -Office $Office `
            -Department $Department `
            -Description $Description `
            -EmployeeID $EmployeeID `

        Write-Host "Usuário "$dadosUsuario.SamAccountName" atualizado com sucesso."
    } catch {
        Write-Error "Erro ao atualizar o usuário "$dadosUsuario.SamAccountName""
    }
}
