<#
.SYNOPSIS
    Substitui a imagem de papel de parede por uma nova imagem e, opcionalmente, define o estilo.
.PARAMETER NovaImagem
    O caminho completo para a nova imagem que será usada como papel de parede.
.EXAMPLE
    .\TrocarWallpaper.ps1 -NovaImagem "C:\Users\Public\Pictures\Sample Pictures\Koala.jpg"
#>
param(
    [Parameter(Mandatory=$true)]
    [string]$NovaImagem,

    [Parameter(Mandatory=$false)]
    [ValidateSet('Centralizar', 'Repetir', 'Estender', 'Abranger')]
    [string]$Estilo
)

# Definições da interface COM IDesktopWallpaper
$comCode = @"
using System;
using System.Runtime.InteropServices;

// Define the IDesktopWallpaper interface
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
// Helper managed que instancia a coclass Desktop Wallpaper e chama a interface IDesktopWallpaper
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
# Garante que os tipos COM não sejam definidos mais de uma vez na mesma sessão
if (-not ([System.Management.Automation.PSTypeName]'IDesktopWallpaper').Type) {
    Add-Type -TypeDefinition $comCode -ErrorAction Stop -PassThru | Out-Null
}


# Verifica se a nova imagem fornecida existe
if (-not (Test-Path $NovaImagem)) {
    Write-Error "Erro: O arquivo de imagem '$NovaImagem' não foi encontrado."
    exit 1
}

# Define o caminho de destino do wallpaper da política
$caminhoDestino = "$env:USERPROFILE\tpu.png"

try {
    # Se o arquivo de destino existir, faz um backup
    if (Test-Path $caminhoDestino) {
        $caminhoBackup = "$caminhoDestino.bak"
        Write-Host "Fazendo backup do papel de parede atual para '$caminhoBackup'..."
        Copy-Item -Path $caminhoDestino -Destination $caminhoBackup -Force -ErrorAction Stop
        Write-Host "Backup concluído."
    }

    # Copia a nova imagem para o destino, substituindo o arquivo existente
    Write-Host "Substituindo o papel de parede em '$caminhoDestino'..."
    Copy-Item -Path $NovaImagem -Destination $caminhoDestino -Force -ErrorAction Stop

    Write-Host ""
    Write-Host "Sucesso! O papel de parede foi substituído."

    # --- Recupera dimensões da imagem aplicada ---
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
        $img = [System.Drawing.Image]::FromFile($caminhoDestino)
        $imgWidth = $img.Width
        $imgHeight = $img.Height
        $img.Dispose()
        Write-Host "Dimensões da imagem aplicada: ${imgWidth}x${imgHeight} (LxA)"
    } catch {
        Write-Warning "Não foi possível ler dimensões da imagem: $($_.Exception.Message)"
        $imgWidth = $null; $imgHeight = $null
    }

    # --- Recupera resolução combinada dos monitores ---
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $screens = [System.Windows.Forms.Screen]::AllScreens
        $totalWidth = 0
        $maxHeight = 0
        $i = 0
        foreach ($s in $screens) {
            $i++
            Write-Host ("Monitor {0}: {1}x{2} @ {3}" -f $i, $s.Bounds.Width, $s.Bounds.Height, $s.DeviceName)
            $totalWidth += $s.Bounds.Width
            if ($s.Bounds.Height -gt $maxHeight) { $maxHeight = $s.Bounds.Height }
        }
        Write-Host "Resolução combinada (soma larguras x maior altura): ${totalWidth}x${maxHeight} (LxA)"
    } catch {
        Write-Warning "Não foi possível obter informações dos monitores: $($_.Exception.Message)"
        $totalWidth = $null; $maxHeight = $null
    }

    # --- Recupera lista de monitores via IDesktopWallpaper (COM) para comparação ---
    try {
        $monInfos = [WallpaperHelper]::GetMonitorsInfo()
        if ($null -ne $monInfos) {
            Write-Host "Monitores reportados pela API IDesktopWallpaper:"
            foreach ($m in $monInfos) {
                # formato: index;path;width;height
                $parts = $m -split ';'
                $idx = $parts[0]; $path = $parts[1]; $w = $parts[2]; $h = $parts[3]
                Write-Host (" Monitor {0}: {1} - {2}x{3}" -f $idx, $path, $w, $h)
            }
        }
    } catch {
        Write-Warning "Não foi possível obter lista de monitores via COM: $($_.Exception.Message)"
    }

    # --- Comparação e veredito de compatibilidade para 'Abranger' (span) ---
    if ($null -ne $imgWidth -and $null -ne $imgHeight -and $null -ne $totalWidth -and $null -ne $maxHeight) {
        if ($imgWidth -eq $totalWidth -and $imgHeight -eq $maxHeight) {
            Write-Host 'Compatibilidade: OK - a imagem tem exatamente a resolucao combinada dos monitores (compativel com span).'
        } else {
            Write-Host 'Compatibilidade: NAO - a imagem NAO tem as dimensoes exatas para span.'
            Write-Host "  - Imagem: ${imgWidth}x${imgHeight}"
            Write-Host "  - Area span esperada: ${totalWidth}x${maxHeight}"
            if ($imgWidth -lt $totalWidth -or $imgHeight -lt $maxHeight) {
                Write-Host '  Observacao: a imagem e menor do que a area span; o Windows pode repetir/centralizar/esticar em vez de span.'
            } else {
                Write-Host '  Observacao: a imagem e maior que a area span; o Windows pode recortar ou centrar dependendo da posicao.' 
            }
        }
    }

    # Tenta configurar o papel de parede e o estilo via COM, que é o método preferido
    Write-Host "Configurando papel de parede e estilo via COM (helper gerenciado)..."
    try {
        $styleMap = @{ 
            'Centralizar' = [DesktopWallpaperPosition]::Center
            'Repetir'     = [DesktopWallpaperPosition]::Tile
            'Estender'    = [DesktopWallpaperPosition]::Stretch
            'Abranger'    = [DesktopWallpaperPosition]::Span
        }

        if ($PSBoundParameters.ContainsKey('Estilo')) {
            $position = $styleMap[$Estilo]
            Write-Host "Aplicando estilo via COM: $Estilo"
            [WallpaperHelper]::SetWallpaperAndPosition($caminhoDestino, $position)
            Write-Host "Estilo '$Estilo' aplicado via COM."
        } else {
            # Apenas define o wallpaper, sem alterar posição
            [WallpaperHelper]::SetWallpaperAndPosition($caminhoDestino, [DesktopWallpaperPosition]::Center)
            Write-Host "Papel de parede definido via COM (posição padrão)."
        }

    } catch {
        Write-Warning "Falha ao usar a interface COM para definir o papel de parede. Tentando método alternativo..."
        Write-Warning $_.Exception.Message
        # Método antigo como fallback: define estilo via registro (se informado) e usa SystemParametersInfo
        if ($PSBoundParameters.ContainsKey('Estilo')) {
            $styleMapReg = @{ 
                'Centralizar' = @{ WallpaperStyle='0'; TileWallpaper='0' }
                'Repetir'     = @{ WallpaperStyle='0'; TileWallpaper='1' }
                'Estender'    = @{ WallpaperStyle='2'; TileWallpaper='0' }
                # 'Abranger' (Span) costuma ser representado por WallpaperStyle=22, TileWallpaper=0
                'Abranger'    = @{ WallpaperStyle='22'; TileWallpaper='0' }
            }
            $regVals = $styleMapReg[$Estilo]
            if ($null -ne $regVals) {
                try {
                    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name 'WallpaperStyle' -Value $regVals.WallpaperStyle -ErrorAction Stop
                    Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name 'TileWallpaper' -Value $regVals.TileWallpaper -ErrorAction Stop
                    Write-Host "Estilo '$Estilo' aplicado via registro."
                } catch {
                    Write-Warning "Falha ao escrever configuração de estilo no registro: $($_.Exception.Message)"
                }
            }
        }
        # Converte para BMP antes de chamar SystemParametersInfo — algumas versões do Windows
        # aplicam corretamente o estilo somente quando o arquivo é BMP.
        $bmpDestino = "$env:USERPROFILE\tpu.bmp"
        try {
            Add-Type -AssemblyName System.Drawing
            $img = [System.Drawing.Image]::FromFile($caminhoDestino)
            $img.Save($bmpDestino, [System.Drawing.Imaging.ImageFormat]::Bmp)
            $img.Dispose()
            $spTarget = $bmpDestino
        } catch {
            Write-Warning "Falha ao converter imagem para BMP: $($_.Exception.Message). Tentando usar arquivo original.";
            $spTarget = $caminhoDestino
        }

        $code = 'using System.Runtime.InteropServices; public class W { [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); }'
        Add-Type $code
        [W]::SystemParametersInfo(20, 0, $spTarget, 3) | Out-Null
    } finally {
        if ($null -ne $wallpaperManager) {
            try {
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wallpaperManager) | Out-Null
            } catch {
                # Ignora erros ao liberar o wrapper COM (p.ex. já liberado)
            }
        }
        if ($null -ne $comObject) {
            try {
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null
            } catch {
                # Ignora erros ao liberar o objeto COM
            }
        }
    }

    # Após aplicar (COM ou fallback), tentar ler posição atual via COM e limpar cache + reiniciar Explorer
    try {
        $pos = $null
        try {
            $pos = [WallpaperHelper]::GetPosition()
            Write-Host "Posição atual reportada pelo COM: $pos"
        } catch {
            Write-Warning "Não foi possível ler posição via COM: $($_.Exception.Message)"
        }

        # Limpa cache de wallpapers que o Explorer pode usar
        $themeDir = Join-Path $env:APPDATA 'Microsoft\Windows\Themes'
        if (Test-Path $themeDir) {
            Get-ChildItem -Path $themeDir -Filter 'Transcoded*' -File -ErrorAction SilentlyContinue | ForEach-Object {
                try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue } catch {}
            }
            Get-ChildItem -Path $themeDir -Filter 'CachedFiles*' -File -ErrorAction SilentlyContinue | ForEach-Object {
                try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue } catch {}
            }
            Write-Host "Cache de papel de parede limpo em '$themeDir'."
        }

        # Reinicia o explorer para forçar recarregamento do wallpaper
        Write-Host "Reiniciando o Explorer para aplicar as mudanças..."
        try {
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        } catch {}
        Start-Process explorer

    } catch {
        Write-Warning "Erro ao tentar limpar cache/reiniciar Explorer: $($_.Exception.Message)"
    }

    Write-Host "Operação concluída."

} catch {
    Write-Error "Ocorreu um erro ao tentar substituir o arquivo:"
    Write-Error $_.Exception.Message
    exit 1
}
