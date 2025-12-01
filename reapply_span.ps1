# Tenta reaplicar SetPosition(Span) via COM varias vezes com reinicio do Explorer entre tentativas
param(
    [int]$Attempts = 3,
    [int]$DelaySeconds = 2
)

# Define helper minimo para SetPosition/GetPosition
$comCode = @"
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IDesktopWallpaper { void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.LPWStr)] string wallpaper); DesktopWallpaperPosition GetPosition(); void SetPosition(DesktopWallpaperPosition position); IntPtr GetSlideshow(); void SetSlideshow(IntPtr items); void Enable(bool enable); }
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
public enum DesktopWallpaperPosition { Center, Tile, Stretch, Fit, Fill, Span }
public enum DesktopSlideshowOptions { ShuffleImages = 0x01 }
public static class WallpaperHelper { public static void SetPositionSpan() { var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var comType = Type.GetTypeFromCLSID(clsid); var comObj = Activator.CreateInstance(comType); var dw = (IDesktopWallpaper)comObj; dw.SetPosition(DesktopWallpaperPosition.Span); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(comObj); }
public static DesktopWallpaperPosition GetPosition(){ var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var comType = Type.GetTypeFromCLSID(clsid); var comObj = Activator.CreateInstance(comType); var dw = (IDesktopWallpaper)comObj; var p = dw.GetPosition(); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(comObj); return p; }
public static bool IsSlideshowActive(){ var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var comType = Type.GetTypeFromCLSID(clsid); var comObj = Activator.CreateInstance(comType); var dw = (IDesktopWallpaper)comObj; var p = dw.GetSlideshow(); bool active = (p != IntPtr.Zero); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(comObj); return active; }
public static void DisableSlideshow(){ var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var comType = Type.GetTypeFromCLSID(clsid); var comObj = Activator.CreateInstance(comType); var dw = (IDesktopWallpaper)comObj; dw.SetSlideshow(IntPtr.Zero); dw.Enable(false); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(comObj);} }
"@

if (-not ([System.Management.Automation.PSTypeName]'IDesktopWallpaper').Type) {
    try { Add-Type -TypeDefinition $comCode -ErrorAction Stop } catch { Write-Warning "Nao foi possivel adicionar helper COM: $($_.Exception.Message)"; exit 1 }
}

for ($i=1; $i -le $Attempts; $i++) {
    Write-Host ("Tentativa {0} de {1}:" -f $i, $Attempts)
    try {
        try {
            $ss = [WallpaperHelper]::IsSlideshowActive()
        } catch { $ss = $false }
        if ($ss) {
            Write-Host " Slideshow ativo -> desabilitando (temporario)"
            try { [WallpaperHelper]::DisableSlideshow() } catch { Write-Warning "Falha ao desabilitar slideshow: $($_.Exception.Message)" }
        }

        [WallpaperHelper]::SetPositionSpan()
        Start-Sleep -Seconds 1
        $pos = [WallpaperHelper]::GetPosition()
        Write-Host "  Posicao apos SetPosition: $pos"
        if ($pos -eq 'Span') { Write-Host '  Resultado: Span aplicado com sucesso.'; break }

        Write-Host '  Nao obteve Span; reiniciando Explorer e repetindo...'
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Process explorer
        Start-Sleep -Seconds $DelaySeconds
        $pos2 = [WallpaperHelper]::GetPosition()
        Write-Host "  Posicao apos restart Explorer: $pos2"
        if ($pos2 -eq 'Span') { Write-Host '  Resultado: Span aplicado apos restart.'; break }
    } catch {
        Write-Warning ("Erro na tentativa {0}: {1}" -f $i, $_.Exception.Message)
    }
    Start-Sleep -Seconds $DelaySeconds
}

# Resultado final
try { $final = [WallpaperHelper]::GetPosition(); Write-Host "Posicao final reportada pelo COM: $final" } catch { Write-Warning "Nao foi possivel ler posicao final: $($_.Exception.Message)" }

# Se ainda nao for Span, listar arquivos de cache para proxima etapa
if ($final -ne 'Span') {
    $themeDir = Join-Path $env:APPDATA 'Microsoft\Windows\Themes'
    if (Test-Path $themeDir) {
        Write-Host 'Arquivos em Themes (para inspecao):'
        Get-ChildItem -Path $themeDir -Force | Format-Table Name,Length,LastWriteTime -AutoSize
    } else { Write-Host 'Themes folder nao encontrado.' }
}
