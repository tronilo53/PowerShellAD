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
        Write-Host "$plus - $nameOU"
    }
    Write-Host ""
    $option = Read-Host "Ingrese Valor"
    $plus = $option - 1;
    $pathOU = $ousRoot[$plus].DistinguishedName
    $nameOrganizational = $ousRoot[$plus].Name
    $ousRoot = @(Get-ADOrganizationalUnit -LDAPFilter '(name=*)' -SearchBase $pathOU -SearchScope OneLevel | Select-Object Name, DistinguishedName)
    while ($ousRoot.Length -gt 0) {
        menuAddUsersSub -path $pathOU -nameOraganizational $nameOrganizational
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
        Write-Host "$plus - $nameOU"
    }
    Write-Host ""
    $option = Read-Host "Ingrese Valor"
    $plus = $option - 1;
    $pathOU = $ousRoot[$plus].DistinguishedName
    $nameOraganizationalSub = $ousRoot[$plus].Name
    $ousRoot = @(Get-ADOrganizationalUnit -LDAPFilter '(name=*)' -SearchBase $pathOU -SearchScope OneLevel | Select-Object Name, DistinguishedName)
    if($ousRoot.Length -gt 0) {
        menuAddUsersSub -path $pathOU -nameOraganizational $nameOraganizationalSub
    }else {
        break
    }
}

#Importarcion de Active Directory
Import-Module ActiveDirectory

menuMain

while (($option = Read-Host -Prompt "Ingrese Valor") -ne "6") {
    switch ($option) {
        1 {
            menuAddUsers
            Pause
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