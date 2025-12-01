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
    [ValidateSet('Centralizar', 'Repetir', 'Estender')]
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

    # Tenta configurar o papel de parede e o estilo via COM, que é o método preferido
    Write-Host "Configurando papel de parede e estilo via COM..."
    $wallpaperManager = $null
    try {
        # Cria a instância do objeto COM DesktopWallpaper usando seu ProgID
        $wallpaperManager = New-Object -ComObject DesktopWallpaper
        
        # Obtém a interface IDesktopWallpaper do objeto COM
        $desktopWallpaper = [Runtime.InteropServices.Marshal]::GetTypedObjectForIUnknown($wallpaperManager, [IDesktopWallpaper])

        # Define o papel de parede para todos os monitores (primeiro argumento nulo)
        $desktopWallpaper.SetWallpaper($null, $caminhoDestino)

        # Se um estilo foi especificado, aplica-o
        if ($PSBoundParameters.ContainsKey('Estilo')) {
             Write-Host "Aplicando estilo '$Estilo'..."
             # Mapeia o nome do estilo para o valor do enum
             $styleMap = @{ 
                 'Centralizar' = [DesktopWallpaperPosition]::Center
                 'Repetir'     = [DesktopWallpaperPosition]::Tile
                 'Estender'    = [DesktopWallpaperPosition]::Stretch
             }
             $position = $styleMap[$Estilo]
 
             # Define a posição (estilo) do papel de parede
             $desktopWallpaper.SetPosition($position)
             Write-Host "Estilo '$Estilo' aplicado com sucesso."
        } else {
            Write-Host "Nenhum estilo especificado. O padrão do sistema será usado."
        }

    } catch {
        Write-Warning "Falha ao usar a interface COM para definir o papel de parede. Tentando método alternativo..."
        Write-Warning $_.Exception.Message
        # Método antigo como fallback
        $code = 'using System.Runtime.InteropServices; public class W { [DllImport("user32.dll", CharSet=CharSet.Auto)] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); }'
        Add-Type $code
        [W]::SystemParametersInfo(20, 0, $caminhoDestino, 3) | Out-Null
    } finally {
        if ($null -ne $wallpaperManager) {
            # Libera o objeto COM
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wallpaperManager) | Out-Null
        }
    }

    Write-Host "Operação concluída."

} catch {
    Write-Error "Ocorreu um erro ao tentar substituir o arquivo:"
    Write-Error $_.Exception.Message
    exit 1
}
