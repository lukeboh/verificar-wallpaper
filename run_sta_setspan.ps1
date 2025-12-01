# Script para ser executado em um PowerShell iniciado com -STA
# Faz Add-Type para IDesktopWallpaper e chama SetPosition(Span) e GetPosition
$comCode = @'
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IDesktopWallpaper { void SetPosition(DesktopWallpaperPosition position); DesktopWallpaperPosition GetPosition(); IntPtr GetSlideshow(); void SetSlideshow(IntPtr items); void Enable(bool enable); }
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
public enum DesktopWallpaperPosition { Center, Tile, Stretch, Fit, Fill, Span }
public static class WallpaperHelper { public static void SetPositionSpan(){ var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var t = Type.GetTypeFromCLSID(clsid); var o = Activator.CreateInstance(t); var dw = (IDesktopWallpaper)o; dw.SetPosition(DesktopWallpaperPosition.Span); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(o); }
public static DesktopWallpaperPosition GetPosition(){ var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var t = Type.GetTypeFromCLSID(clsid); var o = Activator.CreateInstance(t); var dw = (IDesktopWallpaper)o; var p = dw.GetPosition(); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(o); return p; } }
'@

Add-Type -TypeDefinition $comCode -ErrorAction Stop
Write-Host 'Add-Type carregado no processo STA.'
try {
    [WallpaperHelper]::SetPositionSpan()
    Start-Sleep -Milliseconds 500
    $pos = [WallpaperHelper]::GetPosition()
    Write-Host "Posicao após SetPosition (no STA): $pos"
} catch {
    Write-Error "Erro no STA: $($_.Exception.Message)"
    exit 1
}
