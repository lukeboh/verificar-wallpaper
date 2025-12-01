<#
.SYNOPSIS
    Substitui a imagem de papel de parede por uma nova imagem e, opcionalmente, define o estilo.
.PARAMETER NovaImagem
    O caminho completo para a nova imagem que será usada como papel de parede.
.PARAMETER Estilo
    Centralizar | Repetir | Estender | Abranger
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$NovaImagem,

    [Parameter(Mandatory=$false)]
    [ValidateSet('Centralizar','Repetir','Estender','Abranger')]
    [string]$Estilo
)

# C# helper via Add-Type (usado para chamar IDesktopWallpaper COM coclass)
$comCode = @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("B92B56A9-8B55-4E14-9A89-0199BBB6F93B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IDesktopWallpaper 
{
    void SetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [MarshalAs(UnmanagedType.LPWStr)] string wallpaper);
    [return: MarshalAs(UnmanagedType.LPWStr)]
    string GetWallpaper([MarshalAs(UnmanagedType.LPWStr)] string monitorID);
    [return: MarshalAs(UnmanagedType.LPWStr)]
    string GetMonitorDevicePathAt(uint monitorIndex);
    uint GetMonitorDevicePathCount();
    void GetMonitorRECT([MarshalAs(UnmanagedType.LPWStr)] string monitorID, out RECT rect);
    void SetBackgroundColor(uint color);
    uint GetBackgroundColor();
    void SetPosition(DesktopWallpaperPosition position);
    DesktopWallpaperPosition GetPosition();
    void SetSlideshow(IntPtr items);
    IntPtr GetSlideshow();
    void SetSlideshowOptions(DesktopSlideshowOptions options, uint slideshowTick);
    void GetSlideshowOptions(out DesktopSlideshowOptions options, out uint slideshowTick);
    void AdvanceSlideshow([MarshalAs(UnmanagedType.LPWStr)] string monitorID, [In] DesktopSlideshowDirection direction);
    DesktopSlideshowDirection GetStatus();
    void Enable(bool enable);
}

[StructLayout(LayoutKind.Sequential)]
public struct RECT { public int Left, Top, Right, Bottom; }

public enum DesktopWallpaperPosition { Center, Tile, Stretch, Fit, Fill, Span }
public enum DesktopSlideshowDirection { Forward, Backward }
public enum DesktopSlideshowOptions { ShuffleImages = 0x01 }

public static class WallpaperHelper {
    public static void SetWallpaperAndPosition(string wallpaper, DesktopWallpaperPosition position) {
        var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD");
        var comType = Type.GetTypeFromCLSID(clsid);
        var comObj = Activator.CreateInstance(comType);
        var dw = (IDesktopWallpaper)comObj;
        dw.SetWallpaper(null, wallpaper);
        dw.SetPosition(position);
        Marshal.ReleaseComObject(dw);
        Marshal.ReleaseComObject(comObj);
    }

    public static DesktopWallpaperPosition GetPosition() {
        var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD");
        var comType = Type.GetTypeFromCLSID(clsid);
        var comObj = Activator.CreateInstance(comType);
        var dw = (IDesktopWallpaper)comObj;
        var pos = dw.GetPosition();
        Marshal.ReleaseComObject(dw);
        Marshal.ReleaseComObject(comObj);
        return pos;
    }

    public static bool IsSlideshowActive() {
        var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD");
        var comType = Type.GetTypeFromCLSID(clsid);
        var comObj = Activator.CreateInstance(comType);
        var dw = (IDesktopWallpaper)comObj;
        var p = dw.GetSlideshow();
        bool active = (p != IntPtr.Zero);
        Marshal.ReleaseComObject(dw);
        Marshal.ReleaseComObject(comObj);
        return active;
    }

    public static void DisableSlideshow() {
        var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD");
        var comType = Type.GetTypeFromCLSID(clsid);
        var comObj = Activator.CreateInstance(comType);
        var dw = (IDesktopWallpaper)comObj;
        dw.SetSlideshow(IntPtr.Zero);
        dw.Enable(false);
        Marshal.ReleaseComObject(dw);
        Marshal.ReleaseComObject(comObj);
    }

    public static string[] GetMonitorsInfo() {
        var clsid = new Guid("C2CF3110-460E-4fc1-B9D0-8A1C0C9CC4BD");
        var comType = Type.GetTypeFromCLSID(clsid);
        var comObj = Activator.CreateInstance(comType);
        var dw = (IDesktopWallpaper)comObj;
        uint count = dw.GetMonitorDevicePathCount();
        int c = (int)count;
        var arr = new string[c];
        for (int i = 0; i < c; i++) {
            string path = dw.GetMonitorDevicePathAt((uint)i);
            RECT rect;
            dw.GetMonitorRECT(path, out rect);
            arr[i] = string.Format("{0};{1};{2};{3}", i, path, rect.Right - rect.Left, rect.Bottom - rect.Top);
        }
        Marshal.ReleaseComObject(dw);
        Marshal.ReleaseComObject(comObj);
        return arr;
    }
}
"@

# Add the managed helper type
if (-not ([System.Management.Automation.PSTypeName]'IDesktopWallpaper').Type) {
    Add-Type -TypeDefinition $comCode -ErrorAction Stop -PassThru | Out-Null
}

# Verifica se a nova imagem fornecida existe
if (-not (Test-Path $NovaImagem)) {
    Write-Error "Erro: O arquivo de imagem '$NovaImagem' não foi encontrado."
    exit 1
}

# destino do wallpaper
$caminhoDestino = "$env:USERPROFILE\tpu.png"

try {
    if (Test-Path $caminhoDestino) {
        $caminhoBackup = "$caminhoDestino.bak"
        Write-Host "Fazendo backup do papel de parede atual para '$caminhoBackup'..."
        Copy-Item -Path $caminhoDestino -Destination $caminhoBackup -Force -ErrorAction Stop
        Write-Host "Backup concluido."
    }

    Write-Host "Substituindo o papel de parede em '$caminhoDestino'..."
    Copy-Item -Path $NovaImagem -Destination $caminhoDestino -Force -ErrorAction Stop
    Write-Host "Sucesso! A imagem foi copiada."

    # dimensoes da imagem
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
        $img = [System.Drawing.Image]::FromFile($caminhoDestino)
        $imgWidth = $img.Width; $imgHeight = $img.Height
        $img.Dispose()
        Write-Host "Dimensoes da imagem: ${imgWidth}x${imgHeight} (LxA)"
    } catch {
        Write-Warning "Nao foi possivel ler dimensoes da imagem: $($_.Exception.Message)"
        $imgWidth = $null; $imgHeight = $null
    }

    # resolucao combinada via System.Windows.Forms
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $screens = [System.Windows.Forms.Screen]::AllScreens
        $totalWidth = 0; $maxHeight = 0; $i = 0
        foreach ($s in $screens) {
            $i++ | Out-Null
            Write-Host ("Monitor {0}: {1}x{2} @ {3}" -f $i, $s.Bounds.Width, $s.Bounds.Height, $s.DeviceName)
            $totalWidth += $s.Bounds.Width
            if ($s.Bounds.Height -gt $maxHeight) { $maxHeight = $s.Bounds.Height }
        }
        Write-Host "Resolucao combinada: ${totalWidth}x${maxHeight} (LxA)"
    } catch {
        Write-Warning "Nao foi possivel obter informacoes dos monitores: $($_.Exception.Message)"
        $totalWidth = $null; $maxHeight = $null
    }

    # lista de monitores via COM
    try {
        $monInfos = [WallpaperHelper]::GetMonitorsInfo()
        if ($null -ne $monInfos) {
            Write-Host "Monitores (IDesktopWallpaper):"
            foreach ($m in $monInfos) {
                $parts = $m -split ';'
                Write-Host (" Monitor {0}: {1} - {2}x{3}" -f $parts[0], $parts[1], $parts[2], $parts[3])
            }
        }
    } catch {
        Write-Warning "Nao foi possivel obter lista de monitores via COM: $($_.Exception.Message)"
    }

    # comparar imagem vs area span
    if ($null -ne $imgWidth -and $null -ne $imgHeight -and $null -ne $totalWidth -and $null -ne $maxHeight) {
        if ($imgWidth -eq $totalWidth -and $imgHeight -eq $maxHeight) {
            Write-Host 'Compatibilidade: OK - imagem tem resolucao combinada (span).'
        } else {
            Write-Host 'Compatibilidade: NAO - imagem nao tem dimensoes exatas para span.'
            Write-Host "  - Imagem: ${imgWidth}x${imgHeight}"
            Write-Host "  - Area span esperada: ${totalWidth}x${maxHeight}"
        }
    }

    # aplicar via COM: desabilitar slideshow se ativo (sem reativacao)
    Write-Host "Configurando via COM..."
    try {
        $styleMap = @{ 'Centralizar'=[DesktopWallpaperPosition]::Center; 'Repetir'=[DesktopWallpaperPosition]::Tile; 'Estender'=[DesktopWallpaperPosition]::Stretch; 'Abranger'=[DesktopWallpaperPosition]::Span }
        if ($PSBoundParameters.ContainsKey('Estilo')) {
            $position = $styleMap[$Estilo]
            try { $ss = [WallpaperHelper]::IsSlideshowActive() } catch { $ss = $false }
            if ($ss) {
                Write-Host "Slideshow ativo -> desabilitando via COM (sem reativacao)."
                try { [WallpaperHelper]::DisableSlideshow() } catch { Write-Warning "Falha ao desabilitar slideshow: $($_.Exception.Message)" }
            }
            [WallpaperHelper]::SetWallpaperAndPosition($caminhoDestino, $position)
            Write-Host "Wallpaper aplicado via COM (pos: $position)."
        } else {
            [WallpaperHelper]::SetWallpaperAndPosition($caminhoDestino, [DesktopWallpaperPosition]::Center)
            Write-Host "Wallpaper aplicado via COM (pos padrao)."
        }
    } catch {
        Write-Warning "Falha COM: $($_.Exception.Message)"
        # fallback: ajustar registro e usar SystemParametersInfo
        if ($PSBoundParameters.ContainsKey('Estilo')) {
            $styleMapReg = @{ 'Centralizar'=@{WallpaperStyle='0';TileWallpaper='0'}; 'Repetir'=@{WallpaperStyle='0';TileWallpaper='1'}; 'Estender'=@{WallpaperStyle='2';TileWallpaper='0'}; 'Abranger'=@{WallpaperStyle='22';TileWallpaper='0'} }
            $vals = $styleMapReg[$Estilo]
            if ($null -ne $vals) {
                try { Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name 'WallpaperStyle' -Value $vals.WallpaperStyle; Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name 'TileWallpaper' -Value $vals.TileWallpaper } catch { Write-Warning "Falha ao gravar registro: $($_.Exception.Message)" }
            }
        }
        # converter para BMP e chamar SystemParametersInfo
        $bmpDestino = "$env:USERPROFILE\tpu.bmp"
        try { Add-Type -AssemblyName System.Drawing; $img2 = [System.Drawing.Image]::FromFile($caminhoDestino); $img2.Save($bmpDestino,[System.Drawing.Imaging.ImageFormat]::Bmp); $img2.Dispose(); $sp = $bmpDestino } catch { $sp = $caminhoDestino }
        $code = 'using System.Runtime.InteropServices; public class W { [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); }'
        Add-Type $code -ErrorAction SilentlyContinue
        [W]::SystemParametersInfo(20,0,$sp,3) | Out-Null
    }

    # ler posicao atual via COM
    try { $pos = [WallpaperHelper]::GetPosition(); Write-Host "Posicao atual (COM): $pos" } catch { Write-Warning "Nao foi possivel ler posicao via COM: $($_.Exception.Message)" }

    # limpar cache e reiniciar explorer para forcar recarga
    try {
        $themeDir = Join-Path $env:APPDATA 'Microsoft\\Windows\\Themes'
        if (Test-Path $themeDir) {
            Get-ChildItem -Path $themeDir -Filter 'Transcoded*' -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path $themeDir -Filter 'CachedFiles*' -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            Write-Host "Cache limpo em $themeDir"
        }
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Process explorer
    } catch { Write-Warning "Erro ao reiniciar explorer: $($_.Exception.Message)" }

    Write-Host "Operacao concluida."

} catch {
    Write-Error "Erro ao substituir o arquivo: $($_.Exception.Message)"
    exit 1
}
