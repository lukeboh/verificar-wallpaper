# Aplica forçadamente o modo Span: define WallpaperStyle=22, limpa histórico/cache, aplica BMP via SPI, reinicia Explorer e verifica posição COM
param(
    [string]$SourceImage = 'C:\_prj\git\verificar-wallpaper\montanhas.jpg'
)

function Backup-RegistryKey {
    param($Path, $OutFile)
    try {
        # Converter path tipo PowerShell (HKCU:\...) para formato aceito pelo reg.exe (HKEY_CURRENT_USER\...)
        $regPath = $Path -replace '^HKCU:\\','HKEY_CURRENT_USER\\' -replace '^HKLM:\\','HKEY_LOCAL_MACHINE\\'
        reg export "$regPath" "$OutFile" /y | Out-Null
        Write-Host "Backup do registro exportado para $OutFile"
    } catch {
        Write-Warning ("Falha ao exportar {0}: {1}" -f $Path, $_.Exception.Message)
    }
}

try {
    $themeDir = Join-Path $env:APPDATA 'Microsoft\Windows\Themes'
    $wallKey = 'HKCU:\Control Panel\Desktop'
    $explWallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers'

    Write-Host '--- Backup: exportando chaves relevantes ---'
    $bk1 = Join-Path $env:TEMP 'DesktopKey.reg'
    $bk2 = Join-Path $env:TEMP 'ExplorerWallpapers.reg'
    Backup-RegistryKey 'HKCU\\Control Panel\\Desktop' $bk1
    Backup-RegistryKey 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Wallpapers' $bk2

    Write-Host '--- Mostrando valores atuais (antes) ---'
    Get-ItemProperty -Path $wallKey -Name Wallpaper,WallpaperStyle,TileWallpaper -ErrorAction SilentlyContinue | Format-List
    Get-ItemProperty -Path $explWallKey -ErrorAction SilentlyContinue | Format-List

    Write-Host '--- Ajustando WallpaperStyle para Span (22) ---'
    Set-ItemProperty -Path $wallKey -Name WallpaperStyle -Value '22' -ErrorAction Stop
    Set-ItemProperty -Path $wallKey -Name TileWallpaper -Value '0' -ErrorAction Stop
    Write-Host 'Valores do registro atualizados: WallpaperStyle=22, TileWallpaper=0'

    Write-Host '--- Backup dos BackgroundHistoryPath* e Slideshow entries ---'
    $hist = Get-ItemProperty -Path $explWallKey -ErrorAction SilentlyContinue
    $hist | Out-File -FilePath (Join-Path $env:TEMP 'wallpapers_values.txt') -Force
    Write-Host "Backup salvo em: $env:TEMP\wallpapers_values.txt"

    Write-Host '--- Removendo BackgroundHistoryPath* e SlideshowDirectoryPath1 e definindo SlideshowSourceDirectoriesSet=0 ---'
    $vals = Get-ItemProperty -Path $explWallKey -ErrorAction SilentlyContinue
    if ($null -ne $vals) {
        $props = $vals.PSObject.Properties | Where-Object { $_.Name -like 'BackgroundHistoryPath*' -or $_.Name -like 'SlideshowDirectoryPath*' }
        foreach ($p in $props) {
            try { Remove-ItemProperty -Path $explWallKey -Name $p.Name -ErrorAction Stop; Write-Host "Removed $($p.Name)" } catch { Write-Warning "Could not remove $($p.Name): $($_.Exception.Message)" }
        }
        try { Set-ItemProperty -Path $explWallKey -Name 'SlideshowSourceDirectoriesSet' -Value 0 -ErrorAction Stop; Write-Host 'SlideshowSourceDirectoriesSet set to 0' } catch { Write-Warning 'Could not set SlideshowSourceDirectoriesSet' }
    } else {
        Write-Host 'No Wallpaper Explorer key found; skipping history cleanup.'
    }

    # garantir BMP
    Write-Host '--- Convertendo imagem para BMP (destino: %USERPROFILE%\\tpu_span.bmp) ---'
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    $dstBmp = Join-Path $env:USERPROFILE 'tpu_span.bmp'
    $img = [System.Drawing.Image]::FromFile($SourceImage)
    $img.Save($dstBmp, [System.Drawing.Imaging.ImageFormat]::Bmp)
    $img.Dispose()
    Write-Host "BMP salvo em: $dstBmp"

    # definir Wallpaper (registro) para apontar para o BMP
    Set-ItemProperty -Path $wallKey -Name Wallpaper -Value $dstBmp -ErrorAction SilentlyContinue

    # chama SystemParametersInfo
    Write-Host '--- Chamando SystemParametersInfo para aplicar BMP ---'
    $code = 'using System.Runtime.InteropServices; public class W { [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); }'
    Add-Type $code -ErrorAction SilentlyContinue
    $r = [W]::SystemParametersInfo(20,0,$dstBmp,3)
    Write-Host "SystemParametersInfo retornou: $r"

    # limpar cache de temas
    Write-Host '--- Limpando cache de temas (Transcoded/CachedFiles) ---'
    if (Test-Path $themeDir) {
        Get-ChildItem -Path $themeDir -Filter 'Transcoded*' -File -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "Removendo: $($_.FullName)"; Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
        Get-ChildItem -Path $themeDir -Filter 'CachedFiles*' -File -ErrorAction SilentlyContinue | ForEach-Object { Write-Host "Removendo: $($_.FullName)"; Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    } else { Write-Host "Diretorio de temas nao encontrado: $themeDir" }

    Write-Host '--- Reiniciando Explorer ---'
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Process explorer

    Start-Sleep -Seconds 3

    Write-Host '--- Valores de registro apos aplicacao ---'
    Get-ItemProperty -Path $wallKey -Name Wallpaper,WallpaperStyle,TileWallpaper -ErrorAction SilentlyContinue | Format-List

    Write-Host '--- Consultando posicao via COM helper (WallpaperHelper::GetPosition) ---'
    # Garante que o helper COM esteja definido na sessão antes de consultar
    $comCode = @"
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IDesktopWallpaper { void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.LPWStr)] string wallpaper); string GetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID); string GetMonitorDevicePathAt(uint monitorIndex); uint GetMonitorDevicePathCount(); void GetMonitorRECT([MarshalAs(UnmanagedType.LPWStr)] string monitorID, out RECT rect); void SetBackgroundColor(uint color); uint GetBackgroundColor(); void SetPosition(DesktopWallpaperPosition position); DesktopWallpaperPosition GetPosition(); void SetSlideshow(IntPtr items); IntPtr GetSlideshow(); void SetSlideshowOptions(DesktopSlideshowOptions options, uint slideshowTick); void GetSlideshowOptions(out DesktopSlideshowOptions options, out uint slideshowTick); void AdvanceSlideshow([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [In] DesktopSlideshowDirection direction); DesktopSlideshowDirection GetStatus(); void Enable(bool enable); }
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
public enum DesktopWallpaperPosition { Center, Tile, Stretch, Fit, Fill, Span }
public enum DesktopSlideshowDirection { Forward, Backward }
public enum DesktopSlideshowOptions { ShuffleImages = 0x01 }
public static class WallpaperHelper { public static DesktopWallpaperPosition GetPosition() { var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var comType = Type.GetTypeFromCLSID(clsid); var comObj = Activator.CreateInstance(comType); var dw = (IDesktopWallpaper)comObj; var pos = dw.GetPosition(); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(comObj); return pos; } }
"@
    if (-not ([System.Management.Automation.PSTypeName]'IDesktopWallpaper').Type) {
        try { Add-Type -TypeDefinition $comCode -ErrorAction Stop } catch { Write-Warning "Nao foi possivel adicionar o helper COM: $($_.Exception.Message)" }
    }
    try {
        $pos = [WallpaperHelper]::GetPosition()
        Write-Host "Posicao reportada pelo COM: $pos"
    } catch {
        Write-Warning "Nao foi possivel consultar posicao via COM: $($_.Exception.Message)"
    }

    Write-Host '--- Fim da rotina forçada ---'
} catch {
    Write-Error "Erro: $($_.Exception.Message)"
    exit 1
}