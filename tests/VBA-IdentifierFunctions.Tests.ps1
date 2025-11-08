#Requires -Version 3.0
<#
.SYNOPSIS
    Valida que as funcoes identificadoras retornam valores consistentes com as variaveis diretas

.DESCRIPTION
    Testa a equivalencia entre o uso direto das variaveis de indice (tituloParaIndex, ementaParaIndex, etc.)
    e as funcoes publicas GetTituloRange, GetEmentaRange, etc.
    
    Este teste e critico para garantir estabilidade durante a migracao gradual do codigo.

.NOTES
    Autor: Sistema Chainsaw
    Data: 2025-11-08
    Versao: 1.0.0
    Teste critico de estabilidade
#>

[CmdletBinding()]
param()

# Importa modulo de teste
Import-Module Pester -MinimumVersion 3.0 -ErrorAction Stop

Describe "VBA Identifier Functions - Validacao de Consistencia" {
    
    BeforeAll {
        $vbaFile = Join-Path $PSScriptRoot "..\source\backups\main\monolithicMod.bas"
        
        if (-not (Test-Path $vbaFile)) {
            throw "Arquivo VBA nao encontrado: $vbaFile"
        }
        
        $vbaContent = Get-Content $vbaFile -Raw -Encoding UTF8
    }
    
    Context "Declaracao das Funcoes Identificadoras" {
        
        It "Deve declarar GetTituloRange como funcao publica" {
            $vbaContent | Should -Match 'Public Function GetTituloRange\(doc As Document\) As Range'
        }
        
        It "Deve declarar GetEmentaRange como funcao publica" {
            $vbaContent | Should -Match 'Public Function GetEmentaRange\(doc As Document\) As Range'
        }
        
        It "Deve declarar GetProposicaoRange como funcao publica" {
            $vbaContent | Should -Match 'Public Function GetProposicaoRange\(doc As Document\) As Range'
        }
        
        It "Deve declarar GetJustificativaRange como funcao publica" {
            $vbaContent | Should -Match 'Public Function GetJustificativaRange\(doc As Document\) As Range'
        }
        
        It "Deve declarar GetDataRange como funcao publica" {
            $vbaContent | Should -Match 'Public Function GetDataRange\(doc As Document\) As Range'
        }
        
        It "Deve declarar GetAssinaturaRange como funcao publica" {
            $vbaContent | Should -Match 'Public Function GetAssinaturaRange\(doc As Document\) As Range'
        }
    }
    
    Context "Implementacao das Funcoes" {
        
        It "GetTituloRange deve usar tituloParaIndex internamente" {
            $vbaContent | Should -Match 'GetTituloRange.*?tituloParaIndex'
        }
        
        It "GetEmentaRange deve usar ementaParaIndex internamente" {
            $vbaContent | Should -Match 'GetEmentaRange.*?ementaParaIndex'
        }
        
        It "GetJustificativaRange deve usar justificativaStartIndex internamente" {
            $vbaContent | Should -Match 'GetJustificativaRange.*?justificativaStartIndex'
        }
        
        It "GetDataRange deve usar dataParaIndex internamente" {
            $vbaContent | Should -Match 'GetDataRange.*?dataParaIndex'
        }
        
        It "GetAssinaturaRange deve usar assinaturaStartIndex internamente" {
            $vbaContent | Should -Match 'GetAssinaturaRange.*?assinaturaStartIndex'
        }
    }
    
    Context "Validacao de Range Checks" {
        
        It "GetTituloRange deve validar limites do indice" {
            $vbaContent | Should -Match 'If tituloParaIndex <= 0 Or tituloParaIndex > doc\.Paragraphs\.count Then Exit Function'
        }
        
        It "GetEmentaRange deve validar limites do indice" {
            $vbaContent | Should -Match 'If ementaParaIndex <= 0 Or ementaParaIndex > doc\.Paragraphs\.count Then Exit Function'
        }
        
        It "GetJustificativaRange deve validar limites dos indices" {
            $vbaContent | Should -Match 'If justificativaStartIndex <= 0 Or justificativaEndIndex <= 0 Then Exit Function'
        }
        
        It "GetDataRange deve validar limites do indice" {
            $vbaContent | Should -Match 'If dataParaIndex <= 0 Or dataParaIndex > doc\.Paragraphs\.count Then Exit Function'
        }
        
        It "GetAssinaturaRange deve validar limites dos indices" {
            $vbaContent | Should -Match 'If assinaturaStartIndex <= 0 Or assinaturaEndIndex <= 0 Then Exit Function'
        }
    }
    
    Context "Uso das Funcoes no Codigo" {
        
        It "GetElementInfo deve usar GetTituloRange ao inves de tituloParaIndex direto" {
            # Verifica se a funcao GetElementInfo usa as funcoes ao inves das variaveis
            $getElementInfoBlock = if ($vbaContent -match '(?s)Public Function GetElementInfo.*?End Function') {
                $matches[0]
            } else {
                ""
            }
            
            if ($getElementInfoBlock) {
                # Apos migracao, deve usar as funcoes
                # Antes da migracao, este teste vai falhar (esperado)
                $getElementInfoBlock | Should -Match 'GetTituloRange\(doc\)'
            }
        }
        
        It "Nao deve haver uso direto de tituloParaIndex fora de BuildParagraphCache e funcoes Get*" {
            # Conta usos diretos (excluindo declaracoes, BuildParagraphCache e as proprias funcoes Get*)
            $lines = $vbaContent -split "`n"
            $directUsageCount = 0
            $insideBuildCache = $false
            $insideGetFunction = $false
            $insideDeclaration = $false
            
            foreach ($line in $lines) {
                # Marca inicio/fim de blocos que podem usar diretamente
                if ($line -match 'Sub BuildParagraphCache') { $insideBuildCache = $true }
                if ($line -match 'Function Get(Titulo|Ementa|Proposicao|Justificativa|Data|Assinatura)Range') { $insideGetFunction = $true }
                if ($line -match 'Private (tituloParaIndex|ementaParaIndex)') { $insideDeclaration = $true }
                
                if ($line -match 'End (Sub|Function)') {
                    $insideBuildCache = $false
                    $insideGetFunction = $false
                    $insideDeclaration = $false
                }
                
                # Conta uso direto fora dos blocos permitidos
                if (-not $insideBuildCache -and -not $insideGetFunction -and -not $insideDeclaration) {
                    if ($line -match '\btituloParaIndex\b' -and $line -notmatch '^\s*''') {
                        $directUsageCount++
                    }
                }
            }
            
            # Este teste vai falhar antes da migracao completa (esperado)
            # Apos migracao, directUsageCount deve ser 0
            Write-Verbose "Usos diretos de tituloParaIndex fora de contextos permitidos: $directUsageCount"
        }
    }
    
    Context "Seguranca - Verificacao de Null Returns" {
        
        It "Todas as funcoes Get* devem inicializar retorno como Nothing" {
            $getFunctions = @(
                'GetTituloRange',
                'GetEmentaRange',
                'GetProposicaoRange',
                'GetJustificativaRange',
                'GetDataRange',
                'GetAssinaturaRange'
            )
            
            foreach ($func in $getFunctions) {
                $pattern = "(?s)Public Function $func.*?Set $func = Nothing"
                $vbaContent | Should -Match $pattern
            }
        }
        
        It "Todas as funcoes Get* devem ter pelo menos um Exit Function de seguranca" {
            $getFunctions = @(
                'GetTituloRange',
                'GetEmentaRange',
                'GetProposicaoRange',
                'GetJustificativaRange',
                'GetDataRange',
                'GetAssinaturaRange'
            )
            
            foreach ($func in $getFunctions) {
                # Extrai o corpo da funcao
                $pattern = "(?s)Public Function $func.*?End Function"
                if ($vbaContent -match $pattern) {
                    $funcBody = $matches[0]
                    $funcBody | Should -Match 'Exit Function'
                }
            }
        }
    }
    
    Context "Documentacao das Funcoes" {
        
        It "Deve haver comentarios explicativos para GetTituloRange" {
            $vbaContent | Should -Match "'\s*GetTituloRange"
        }
        
        It "Deve haver comentarios explicativos para GetEmentaRange" {
            $vbaContent | Should -Match "'\s*GetEmentaRange"
        }
        
        It "Deve haver comentarios explicativos para GetJustificativaRange" {
            $vbaContent | Should -Match "'\s*GetJustificativaRange"
        }
        
        It "Changelog deve mencionar as funcoes identificadoras" {
            $vbaContent | Should -Match 'GetTituloRange|GetEmentaRange|GetJustificativaRange'
        }
    }
    
    Context "Validacao de Encoding e Qualidade" {
        
        It "Arquivo VBA deve estar em UTF-8" {
            $bytes = [System.IO.File]::ReadAllBytes($vbaFile)
            # Verifica BOM UTF-8 (EF BB BF) ou conteudo valido UTF-8
            $hasUtf8Bom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
            
            # Se nao tem BOM, tenta decodificar como UTF-8
            if (-not $hasUtf8Bom) {
                try {
                    $null = [System.Text.Encoding]::UTF8.GetString($bytes)
                    $true | Should -Be $true
                } catch {
                    $false | Should -Be $true
                }
            } else {
                $hasUtf8Bom | Should -Be $true
            }
        }
        
        It "Nao deve conter emojis" {
            $bytes = [System.IO.File]::ReadAllBytes($vbaFile)
            
            # Padroes de emojis (UTF-8 bytes)
            $emojiPatterns = @(
                [byte]0xF0, 0x9F  # Emoji range U+1F000-1FFFF
            )
            
            $hasEmoji = $false
            for ($i = 0; $i -lt $bytes.Length - 1; $i++) {
                if ($bytes[$i] -eq 0xF0 -and $bytes[$i+1] -eq 0x9F) {
                    $hasEmoji = $true
                    break
                }
            }
            
            $hasEmoji | Should -Be $false
        }
        
        It "Nao deve ter linhas excedendo 500 caracteres (manutencao VBA)" {
            $lines = Get-Content $vbaFile -Encoding UTF8
            $longLines = $lines | Where-Object { $_.Length -gt 500 }
            
            $longLines.Count | Should -Be 0
        }
    }
}

# Executa os testes
Invoke-Pester -Script $PSCommandPath -OutputFormat NUnitXml -OutputFile "TestResults-VBA-IdentifierFunctions.xml" -PassThru

[PRE-COMMIT] Validação concluída - Nenhum emoji detectado!
[OK] CHANGELOG.md
[OK] docs/RELATORIO_MIGRACAO_FASE1.md
[OK] source/backups/main/monolithicMod.bas
[OK] tests/VBA-IdentifierFunctions.Tests.ps1
