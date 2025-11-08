$filesToCheck = @(
    'c:\Users\csantos\chainsaw\docs\LGPD_CONFORMIDADE.md',
    'c:\Users\csantos\chainsaw\docs\NOVIDADES_v1.1.md',
    'c:\Users\csantos\chainsaw\docs\SEGURANCA_PRIVACIDADE.md',
    'c:\Users\csantos\chainsaw\docs\SEM_PRIVILEGIOS_ADMIN.md',
    'c:\Users\csantos\chainsaw\docs\VALIDACAO_TIPO_DOCUMENTO.md',
    'c:\Users\csantos\chainsaw\LGPD_ATESTADO.md',
    'c:\Users\csantos\chainsaw\README.md',
    'c:\Users\csantos\chainsaw\installation\inst_docs\GUIA_INSTALACAO.md'
)

foreach ($filePath in $filesToCheck) {
    if (-not (Test-Path $filePath)) { continue }
    
    Write-Host "`n=== $([System.IO.Path]::GetFileName($filePath)) ===" -ForegroundColor Cyan
    
    $bytes = [System.IO.File]::ReadAllBytes($filePath)
    $found = $false
    
    for ($i = 0; $i -lt $bytes.Length - 3; $i++) {
        # Emojis F0 9F
        if ($bytes[$i] -eq 0xF0 -and $bytes[$i+1] -eq 0x9F) {
            $b3 = [Convert]::ToString($bytes[$i+2], 16).PadLeft(2, '0')
            $b4 = [Convert]::ToString($bytes[$i+3], 16).PadLeft(2, '0')
            
            # Calcula Unicode codepoint
            $codepoint = (($bytes[$i] -band 0x07) -shl 18) -bor 
                        (($bytes[$i+1] -band 0x3F) -shl 12) -bor
                        (($bytes[$i+2] -band 0x3F) -shl 6) -bor
                        ($bytes[$i+3] -band 0x3F)
            
            $char = [char]::ConvertFromUtf32($codepoint)
            
            Write-Host "  Emoji em pos $i : F0 9F $b3 $b4 => U+$($codepoint.ToString('X4')) $char" -ForegroundColor Yellow
            $found = $true
        }
        
        # Simbolos E2
        if ($bytes[$i] -eq 0xE2) {
            $b2 = $bytes[$i+1]
            $b3 = $bytes[$i+2]
            
            if ($b2 -ge 0x9C -and $b2 -le 0xAD) {
                $codepoint = (($bytes[$i] -band 0x0F) -shl 12) -bor
                            (($bytes[$i+1] -band 0x3F) -shl 6) -bor
                            ($bytes[$i+2] -band 0x3F)
                
                $char = [char]::ConvertFromUtf32($codepoint)
                
                Write-Host "  Simbolo em pos $i : E2 $([Convert]::ToString($b2, 16)) $([Convert]::ToString($b3, 16)) => U+$($codepoint.ToString('X4')) $char" -ForegroundColor Yellow
                $found = $true
            }
        }
    }
    
    if (-not $found) {
        Write-Host "  Nenhum emoji encontrado" -ForegroundColor Green
    }
}
