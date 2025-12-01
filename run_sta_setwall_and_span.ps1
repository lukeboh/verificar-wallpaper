param(
    [string]$WallpaperPath = (Get-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name Wallpaper -ErrorAction SilentlyContinue).Wallpaper
)
if (-not $WallpaperPath) { Write-Error 'Nao foi possivel obter Wallpaper atual do registro.'; exit 1 }

$comCode = @'
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IDesktopWallpaper { void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.LPWStr)] string wallpaper); DesktopWallpaperPosition GetPosition(); void SetPosition(DesktopWallpaperPosition position); IntPtr GetSlideshow(); void SetSlideshow(IntPtr items); void Enable(bool enable); }
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
public enum DesktopWallpaperPosition { Center, Tile, Stretch, Fit, Fill, Span }
public static class WallpaperHelper { public static void SetWallpaperAndSpan(string wallpaper){ var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var t = Type.GetTypeFromCLSID(clsid); var o = Activator.CreateInstance(t); var dw = (IDesktopWallpaper)o; dw.SetWallpaper(null, wallpaper); dw.SetPosition(DesktopWallpaperPosition.Span); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(o); }
public static DesktopWallpaperPosition GetPosition(){ var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var t = Type.GetTypeFromCLSID(clsid); var o = Activator.CreateInstance(t); var dw = (IDesktopWallpaper)o; var p = dw.GetPosition(); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(o); return p; } }
'@

Add-Type -TypeDefinition $comCode -ErrorAction Stop
Write-Host "Executando SetWallpaper(null, $WallpaperPath) e SetPosition(Span) no processo STA..."
try {
    [WallpaperHelper]::SetWallpaperAndSpan($WallpaperPath)
    Start-Sleep -Milliseconds 500
    $p = [WallpaperHelper]::GetPosition()
    Write-Host "Posicao apos operacao (STA): $p"
} catch {
    Write-Error "Erro no STA: $($_.Exception.Message)"
    exit 1
}
