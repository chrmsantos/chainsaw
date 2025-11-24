================================================================================
CHAINSAW - BACKUP COMPLETO
================================================================================

Data do Backup: 2025-11-24 18:28:00
Usuário: csantos
Computador: NOTEPROCLEG07

Este backup contém uma cópia completa das configurações do Word antes da
instalação do CHAINSAW.

================================================================================
CONTEÚDO DO BACKUP
================================================================================

[Templates]
  - Pasta completa: Templates\
  - Tamanho: 0.21 MB
  - Arquivos: 2
  - Inclui: Normal.dotm, Building Blocks, Temas, Estilos, etc.

[Stamp]
  - Arquivo: stamp.png
  - Tamanho: 54.36 KB

[Personalizações]
  - Pasta: Customizations\
  - Arquivos UI: 2
  - Inclui: Ribbon customizations, Quick Access Toolbar, etc.

================================================================================
RESTAURAÇÃO
================================================================================

Para restaurar este backup, execute:

    .\restore-backup.cmd

Ou manualmente:

    powershell.exe -ExecutionPolicy Bypass -File restore-backup.ps1 -BackupPath "C:\Users\csantos\CHAINSAW\backups\full_backup_20251124_182800"

================================================================================
IMPORTANTE
================================================================================

- Este backup é criado AUTOMATICAMENTE antes de cada instalação
- Backups antigos são mantidos por segurança
- Para liberar espaço, remova manualmente backups antigos desta pasta:
  C:\Users\csantos\CHAINSAW\backups

================================================================================
