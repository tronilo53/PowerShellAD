#Importarcion de Active Directory
Import-Module ActiveDirectory

$global:pathSelected = ""
$global:pathSelectedFormat = ""
$global:nameOrganizational = ""

#Menu Principal
function menuMain {
    Clear-Host
    Write-Host "=================================================="
    Write-Host "|            PowerShellAD v1.0.0                 |"
    Write-Host "=================================================="
    Write-Host "|                                                |"
    Write-Host "|            Seleccione Una Opcion:              |"
    Write-Host "|                                                |"
    Write-Host "|            1) Crear Usuarios                   |"
    Write-Host "|            2) Modificar Usuarios               |"
    Write-Host "|            3) Mover Usuarios                   |"
    Write-Host "|            4) Eliminar Usuarios                |"
    Write-Host "|            5) Crear OU's                       |"
    Write-Host "|            6) Salir                            |"
    Write-Host "|                                                |"
    Write-Host "=================================================="
}
#Menu OU's
function menuAddUsers {
    Clear-Host
    Write-Host "=================================================="
    Write-Host "|                Crear Usuarios                  |"
    Write-Host "=================================================="
    Write-Host ""
    Write-Host "Seleccione una Unidad Organizativa Raiz:"
    Write-Host ""

    $ousRoot = Get-ADOrganizationalUnit -LDAPFilter '(name=*)' -SearchBase "OU=Usuarios Henry Schein SPAIN,OU=ES,DC=eu,DC=hsi,DC=local" -SearchScope OneLevel | Select-Object Name, DistinguishedName

    for ($i = 0; $i -lt $ousRoot.Length; $i++) {
        $plus = $i + 1
        $nameOU = $ousRoot[$i].Name
        Write-Host "$plus) $nameOU"
    }
    $plusPlas = $plus + 1
    Write-Host "$plusPlas) Volver al Menu"
    Write-Host ""
    $option = Read-Host "Ingrese Valor"
    if($option -eq $plusPlas) {
        menuMain
    }else {
        $plus = $option - 1;
        $pathOU = $ousRoot[$plus].DistinguishedName
        $nameOrganizational = $ousRoot[$plus].Name
        $ousRoot = @(Get-ADOrganizationalUnit -LDAPFilter '(name=*)' -SearchBase $pathOU -SearchScope OneLevel | Select-Object Name, DistinguishedName)
        while ($ousRoot.Length -gt 0) {
            $global:pathSelected = $pathOU
            $global:nameOrganizational = $nameOrganizational
            menuAddUsersSub -path $pathOU -nameOraganizational $nameOrganizational
        }
    }
}
function menuAddUsersSub {
    param ($path, $nameOraganizational)
    Clear-Host
    Write-Host "=================================================="
    Write-Host "|                Crear Usuarios                  |"
    Write-Host "=================================================="
    Write-Host ""
    Write-Host "Seleccione una Unidad Organizativa de '" $nameOraganizational "'"
    Write-Host ""

    $ousRoot = @(Get-ADOrganizationalUnit -LDAPFilter '(name=*)' -SearchBase $path -SearchScope OneLevel | Select-Object Name, DistinguishedName)

    for ($i = 0; $i -lt $ousRoot.Length; $i++) {
        $plus = $i + 1
        $nameOU = $ousRoot[$i].Name
        Write-Host "$plus) $nameOU"
    }
    $plusPlas = $plus + 1
    Write-Host "$plusPlas) Volver Al Menu"
    Write-Host ""
    $option = Read-Host "Ingrese Valor"
    if($option -eq $plusPlas) {
        break
        menuMain
    }else {
        $plus = $option - 1;
        $pathOU = $ousRoot[$plus].DistinguishedName
        $nameOraganizationalSub = $ousRoot[$plus].Name
        $ousRoot = @(Get-ADOrganizationalUnit -LDAPFilter '(name=*)' -SearchBase $pathOU -SearchScope OneLevel | Select-Object Name, DistinguishedName)
        if($ousRoot.Length -gt 0) {
            $global:pathSelected = $pathOU
            $global:nameOrganizational = $nameOraganizationalSub
            menuAddUsersSub -path $pathOU -nameOraganizational $nameOraganizationalSub
        }else {
            $global:pathSelected = $pathOU
            $global:nameOrganizational = $nameOraganizationalSub
            break
        }
    }
}
function selectOU {
    param ($pathOU, $nameOU)
    Clear-Host
    Write-Host "=================================================="
    Write-Host "|                Crear Usuarios                  |"
    Write-Host "=================================================="
    Write-Host ""
    Write-Host "Ha Seleccionado la Unidad Organizativa '"$nameOU"'"
    Write-Host ""
    $newPathFile = @("Ous")
    $canonical = @(Get-ADOrganizationalUnit -Properties CanonicalName -Filter * | Sort-Object CanonicalName | Select-Object CanonicalName, DistinguishedName)
    foreach($can in $canonical) {
        if($can.DistinguishedName -eq $pathOU) {
            $pathFile = @($can.CanonicalName -split "/")
            foreach($path in $pathFile) {
                if($path -ne "eu.hsi.local" -and $path -ne "ES" -and $path -ne "Usuarios Henry Schein SPAIN") {
                    $newPathFile = $newPathFile += $path
                }
            }
            Write-Host ""
            $joinedNewPath = $newPathFile -join "/"
            $pathComplete = $joinedNewPath + "/users.csv"
            Write-Host "*Se ejecutara el archivo de la ruta: '"$pathComplete"'"
            Write-Host ""
        }
    }
    $response = Read-Host "Quiere continuar con la ejecucion del script? [S=Si N=No]"
    switch ($response) {
        'S' {
            Write-Host "Creando Usuarios..."
            addUsers -path $pathComplete -pathOU $pathOU
            pause
            break
        }
        'N' {
            Write-Host "Saliendo al menu principal"
            pause
            break
        }
        Default {
            Write-Host -ForegroundColor red -BackgroundColor white "Valor no valido, Por favor, ingrese un valor valido"
            Pause
            selectOU -pathOU $pathOU -nameOU $nameOU
        }
    }
}
function addUsers {
    param ($path, $pathOU)
    $pathNew = "./" + $path
    $users = Import-Csv $pathNew -Delimiter ";"
    foreach($user in $users) {
        $samAccountName = $user.nombre + "." + $user.apellidos
        if(Get-ADuser -Filter {SamAccountName -eq $samAccountName}) {
            Write-Warning "La cuenta '"$samAccountName"' Ya existe y no se creara"
        }else {
            try {
                $userPrincipalName = $user.SamAccountName + "@" + $user.Upn
                $uname = $user.apellidos + ", " + $user.nombre
                $manager = Get-ADuser -Identity $user.manager | Select-Object DistinguishedName
                $object = @{
                    Name = $uname
                    DisplayName = $uname
                    GivenName = $user.nombre
                    Surname = $user.apellidos
                    Department = $user.departamento
                    Title = $user.puestoTrabajo
                    UserPrincipalName = $userPrincipalName
                    SamAccountName = $samAccountName
                    Company = $user.compania
                    Office = $user.oficina
                    Manager = $manager[0].DistinguishedName
                    Path = $pathOU
                    EmailAddress = $userPrincipalName
                    Description = $user.puestoTrabajo
                    AccountPassword = $user.contrasena
                }
                <#
                New-ADUser -Name $uname `
                -DisplayName $uname `
                -GivenName $user.nombre `
                -Surname $user.apellidos `
                -Department $user.departamento `
                -Title $user.puestoTrabajo `
                -UserPrincipalName $userPrincipalName `
                -SamAccountName $samAccountName `
                -Company $user.compania `
                -Office $user.oficina `
                -Manager $manager[0].DistinguishedName `
                -Path $pathOU `
                -EmailAddress $userPrincipalName `
                -Description $user.puestoTrabajo `
                -AccountPassword (ConvertTo-SecureString $user.contrasena -AsPlainText -Force) -Enabled $true -ChangePasswordAtLogon $false -PasswordNeverExpires $false -Verbose
                Set-ADuser -Identity $samAccountName -Replace @{'extensionAttribute10'='EMEAadministration'}
                Write-Host -ForegroundColor Green -BackgroundColor White "Usuario Creado correctamente"
                #>
                Write-Warning $object
            }
            catch {
                Write-Warning "Error al crear el usuario: '"$samAccountName"' Error: "$error[0]
            }
        }
    }
}

menuMain

while (($option = Read-Host -Prompt "Ingrese Valor") -ne "6") {
    switch ($option) {
        1 {
            menuAddUsers
            selectOU -pathOU $global:pathSelected -nameOU $global:nameOrganizational
            break
        }
        2 {
            Clear-Host
            Write-Host "Modificar Usuarios"
            Pause
            break
        }
        3 {
            Clear-Host
            Write-Host "Mover Usuarios"
            Pause
            break
        }
        4 {
            Clear-Host
            Write-Host "Eliminar Usuarios"
            Pause
            break
        }
        5 {
            Clear-Host
            Write-Host "Crear OU's"
            Pause
            break
        }
        6 {
            "exit"
            break
        }
        Default {
            Write-Host -ForegroundColor red -BackgroundColor white "Valor no valido, Por favor, ingrese un valor valido"
            Pause
        }
    }
    menuMain
}