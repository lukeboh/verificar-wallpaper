# Script temporario para converter imagem em BMP e aplicar via SystemParametersInfo
try {
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $src = 'C:\_prj\git\verificar-wallpaper\montanhas.jpg'
    $dst = Join-Path $env:USERPROFILE 'tpu_forced.bmp'
    $img = [System.Drawing.Image]::FromFile($src)
    $img.Save($dst, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $img.Dispose()
    Write-Host "BMP salvo em: $dst"

    $code = 'using System.Runtime.InteropServices; public class W { [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); }'
    Add-Type $code -ErrorAction SilentlyContinue
    $res = [W]::SystemParametersInfo(20,0,$dst,3)
    Write-Host "SystemParametersInfo retornou: $res"

    Write-Host "--- Valores de registro apos fallback ---"
    Get-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper,WallpaperStyle,TileWallpaper -ErrorAction SilentlyContinue | Format-List
} catch {
    Write-Error "Erro no script temporario: $($_.Exception.Message)"
    exit 1
}