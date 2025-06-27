# Importa Módulo ActiveDirectory
Import-Module ActiveDirectory

$excel = Read-Host "Cole (botão direito) o caminho do .csv aqui"
$dadosExcel = Import-Csv -Path $excel -Delimiter ";" -Encoding UTF8

foreach ($usuario in $dadosExcel) {
    $DisplayName = $usuario.NomeCompleto
    $Title = $usuario.Cargo
    $Office = $usuario.Unidade
    $Department = $usuario.Departamento
    $Description = $usuario.CentroCusto
    $employeeNumber = $usuario.Cpf.Trim()
    $employeeID = $usuario.Matricula
    $matriculaGestor = $usuario.MatriculaGestor.Trim()

    # Busca o usuário pelo CPF
    $dadosUsuario = $null
    if (![string]::IsNullOrWhiteSpace($employeeNumber)) {
        try {
            $dadosUsuario = Get-ADUser -Filter "employeeNumber -eq '$employeeNumber'" -Properties *
        } catch {
            Write-Warning "Erro ao buscar usuário com CPF $employeeNumber $_"
            continue
        }

        if (-not $dadosUsuario) {
            Write-Warning "Usuário com CPF $employeeNumber não encontrado no AD. Pulando..."
            continue
        }
    } else {
        Write-Warning "CPF ausente para linha com nome $DisplayName. Pulando..."
        continue
    }

    # Verifica se o nome bate entre CSV e AD
    if ($dadosUsuario.Name -ne $DisplayName) {
        Write-Warning "Nome no AD ($($dadosUsuario.Name)) não bate com CSV ($DisplayName) para CPF $employeeNumber. Pulando por segurança."
        continue
    }

    # Define o gestor do usuário, se informado
    if (![string]::IsNullOrWhiteSpace($matriculaGestor)) {
        try {
            $gestor = Get-ADUser -Filter "employeeID -eq '$matriculaGestor'" -Properties *
            if ($gestor) {
                Set-ADUser -Identity $dadosUsuario.SamAccountName -Manager $gestor.DistinguishedName
            } else {
                Write-Warning "Gestor com matrícula $matriculaGestor não encontrado para o usuário $($dadosUsuario.SamAccountName)."
            }
        } catch {
            Write-Warning "Erro ao buscar ou definir gestor com matrícula $matriculaGestor $_"
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
            -EmployeeID $employeeID

        Write-Host "Usuário $($dadosUsuario.SamAccountName) atualizado com sucesso."
    } catch {
        Write-Error "Erro ao atualizar o usuário $($dadosUsuario.SamAccountName): $_"
    }
}
