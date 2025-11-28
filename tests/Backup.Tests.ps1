# =============================================================================
# Testes do Sistema de Backup Automático (Pester v3 compatible)
# =============================================================================

Describe "Sistema de Backup Automático" {

    BeforeAll {
        # Importa funções de backup
        $script:backupFunctionsPath = Join-Path $PSScriptRoot "..\installation\inst_scripts\backup-functions.ps1"
        if (Test-Path $script:backupFunctionsPath) {
            . $script:backupFunctionsPath
        }
        else {
            throw "Arquivo backup-functions.ps1 não encontrado em $script:backupFunctionsPath"
        }
    }

    Context "Funções de backup disponíveis" {
        
        It "Backup-ChainsawFolder existe" {
            $command = Get-Command Backup-ChainsawFolder -ErrorAction SilentlyContinue
            $command | Should Not BeNullOrEmpty
        }
        
        It "Restore-ChainsawFromBackup existe" {
            $command = Get-Command Restore-ChainsawFromBackup -ErrorAction SilentlyContinue
            $command | Should Not BeNullOrEmpty
        }
        
        It "Remove-ChainsawBackups existe" {
            $command = Get-Command Remove-ChainsawBackups -ErrorAction SilentlyContinue
            $command | Should Not BeNullOrEmpty
        }
    }
    
    Context "Parâmetros das funções" {
        
        It "Backup-ChainsawFolder aceita parâmetro Force" {
            $params = (Get-Command Backup-ChainsawFolder).Parameters
            $params.ContainsKey('Force') | Should Be $true
        }
        
        It "Restore-ChainsawFromBackup aceita parâmetro Force" {
            $params = (Get-Command Restore-ChainsawFromBackup).Parameters
            $params.ContainsKey('Force') | Should Be $true
        }
        
        It "Remove-ChainsawBackups aceita parâmetro KeepLatest" {
            $params = (Get-Command Remove-ChainsawBackups).Parameters
            $params.ContainsKey('KeepLatest') | Should Be $true
        }
    }
    
    Context "Integração com install.ps1" {
        
        It "install.ps1 importa backup-functions.ps1" {
            $installPath = Join-Path $PSScriptRoot "..\installation\inst_scripts\install.ps1"
            $content = Get-Content $installPath -Raw
            $content -match 'backup-functions\.ps1' | Should Be $true
        }
        
        It "install.ps1 chama Backup-ChainsawFolder" {
            $installPath = Join-Path $PSScriptRoot "..\installation\inst_scripts\install.ps1"
            $content = Get-Content $installPath -Raw
            $content -match 'Backup-ChainsawFolder' | Should Be $true
        }
    }
    
    Context "Documentação e feedback" {
        
        It "Backup-ChainsawFolder possui help/comentários" {
            $help = Get-Help Backup-ChainsawFolder -ErrorAction SilentlyContinue
            $help | Should Not BeNullOrEmpty
        }
        
        It "Funções fornecem feedback visual (Write-Host)" {
            $content = Get-Content $script:backupFunctionsPath -Raw
            $content -match 'Write-Host' | Should Be $true
        }
        
        It "Funções usam cores para feedback" {
            $content = Get-Content $script:backupFunctionsPath -Raw
            $content -match '-ForegroundColor' | Should Be $true
        }
    }
    
    Context "Estrutura e qualidade do código" {
        
        It "backup-functions.ps1 implementa tratamento de erros" {
            $content = Get-Content $script:backupFunctionsPath -Raw
            ($content -match 'try\s*\{') -or ($content -match '-ErrorAction') | Should Be $true
        }
        
        It "Funções validam caminhos (Test-Path)" {
            $content = Get-Content $script:backupFunctionsPath -Raw
            $content -match 'Test-Path' | Should Be $true
        }
        
        It "Arquivo possui cabeçalho CHAINSAW" {
            $content = Get-Content $script:backupFunctionsPath -Raw
            $content -match 'CHAINSAW|chainsaw' | Should Be $true
        }
    }
}
