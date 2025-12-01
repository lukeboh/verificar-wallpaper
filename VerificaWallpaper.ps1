
# Verifica políticas de wallpaper no registro
Write-Host "Verificando políticas de papel de parede..."

# Função para checar uma chave
function Check-RegistryKey {
    param (
        [string]$Path
    )
    if (Test-Path $Path) {
        Get-ItemProperty -Path $Path | Select-Object *
    } else {
        Write-Host "Nenhuma chave encontrada em $Path"
    }
}

# Caminhos possíveis
$paths = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\System",
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop",
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System"
)

foreach ($path in $paths) {
    Write-Host "`nChecando: $path"
    Check-RegistryKey -Path $path
}

# Verifica se a política NoChangingWallpaper está ativa
$noChange = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\ActiveDesktop" -ErrorAction SilentlyContinue).NoChangingWallpaper
if ($noChange -eq 1) {
    Write-Host "`n⚠️ Política ativa: Não é permitido alterar o papel de parede."
} else {
    Write-Host "`n✅ Nenhuma política encontrada que bloqueie alteração via ActiveDesktop."
}

# Mostra o papel de parede atual
$currentWallpaper = (Get-ItemProperty -Path "HKCU:\Control Panel\Desktop").Wallpaper
Write-Host "`nPapel de parede atual: $currentWallpaper"
