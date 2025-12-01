# Faz backup de Transcoded* e slideshow.ini, remove-os, reinicia Explorer e tenta SetPosition(Span)
try {
    $themeDir = Join-Path $env:APPDATA 'Microsoft\Windows\Themes'
    if (-not (Test-Path $themeDir)) { Write-Host "Themes dir nao encontrado: $themeDir"; exit 0 }

    $backupDir = Join-Path $env:TEMP "themes_backup_$(Get-Date -Format yyyyMMdd_HHmmss)"
    New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
    Write-Host "Backup dir: $backupDir"

    $files = Get-ChildItem -Path $themeDir -Force | Where-Object { $_.Name -like 'Transcoded*' -or $_.Name -like 'CachedFiles*' -or $_.Name -ieq 'slideshow.ini' }
    if ($files.Count -eq 0) { Write-Host 'No Transcoded/CachedFiles/slideshow.ini found.' } else {
        foreach ($f in $files) {
            $dest = Join-Path $backupDir $f.Name
            Copy-Item -Path $f.FullName -Destination $dest -Force
            Write-Host "Backed up: $($f.Name) -> $dest"
            try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; Write-Host "Removed: $($f.Name)" } catch { Write-Warning "Could not remove $($f.Name): $($_.Exception.Message)" }
        }
    }

    Write-Host 'Reiniciando Explorer...'
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Process explorer
    Start-Sleep -Seconds 3

    # helper COM
    $comCode = @"
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IDesktopWallpaper { void SetPosition(DesktopWallpaperPosition position); DesktopWallpaperPosition GetPosition(); IntPtr GetSlideshow(); void SetSlideshow(IntPtr items); void Enable(bool enable); }
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
public enum DesktopWallpaperPosition { Center, Tile, Stretch, Fit, Fill, Span }
public static class WallpaperHelper { public static void SetPositionSpan(){ var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var t = Type.GetTypeFromCLSID(clsid); var o = Activator.CreateInstance(t); var dw = (IDesktopWallpaper)o; dw.SetPosition(DesktopWallpaperPosition.Span); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(o); }
public static DesktopWallpaperPosition GetPosition(){ var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD"); var t = Type.GetTypeFromCLSID(clsid); var o = Activator.CreateInstance(t); var dw = (IDesktopWallpaper)o; var p = dw.GetPosition(); Marshal.ReleaseComObject(dw); Marshal.ReleaseComObject(o); return p; } }
"@
    if (-not ([System.Management.Automation.PSTypeName]'IDesktopWallpaper').Type) { Add-Type -TypeDefinition $comCode -ErrorAction SilentlyContinue }

    try {
        Write-Host 'Tentando SetPosition(Span) via COM...'
        [WallpaperHelper]::SetPositionSpan()
        Start-Sleep -Milliseconds 500
        $p = [WallpaperHelper]::GetPosition()
        Write-Host "Posicao apos attempt: $p"
    } catch {
        Write-Warning "Erro ao chamar COM: $($_.Exception.Message)"
    }

    Write-Host 'Lista de arquivos atuais em Themes:'
    Get-ChildItem -Path $themeDir -Force | Format-Table Name,Length,LastWriteTime -AutoSize
    Write-Host "Backup mantido em: $backupDir"

} catch {
    Write-Error "Erro no script: $($_.Exception.Message)"
    exit 1
}