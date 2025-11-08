# Converter bytes F0 9F XX YY para caractere Unicode

# F0 9F 94 84 = U+1F504 (counterclockwise arrows button)
# F0 9F 93 9D = U+1F4DD (memo)
# F0 9F 93 A6 = U+1F4E6 (package)
# F0 9F 93 8A = U+1F4CA (bar chart)
# F0 9F 92 BE = U+1F4BE (floppy disk)
# E2 9C A8 = U+2728 (sparkles)

$emojisEncontrados = @(
    @{ Bytes = "F0 9F 94 84"; Unicode = 0x1F504; Nome = "counterclockwise arrows" }
    @{ Bytes = "F0 9F 93 9D"; Unicode = 0x1F4DD; Nome = "memo" }
    @{ Bytes = "F0 9F 93 A6"; Unicode = 0x1F4E6; Nome = "package" }
    @{ Bytes = "F0 9F 93 8A"; Unicode = 0x1F4CA; Nome = "bar chart" }
    @{ Bytes = "F0 9F 92 BE"; Unicode = 0x1F4BE; Nome = "floppy disk" }
    @{ Bytes = "E2 9C A8"; Unicode = 0x2728; Nome = "sparkles" }
)

Write-Host "`nEmojis encontrados no arquivo:`n" -ForegroundColor Cyan

foreach ($emoji in $emojisEncontrados) {
    $char = [char]::ConvertFromUtf32($emoji.Unicode)
    Write-Host "Bytes: $($emoji.Bytes) => $char ($($emoji.Nome))" -ForegroundColor Yellow
    Write-Host "  Unicode: U+$($emoji.Unicode.ToString('X4'))" -ForegroundColor Gray
    Write-Host "  Para adicionar ao mapeamento: " -NoNewline -ForegroundColor Gray
    Write-Host "`$emojiReplacements[[char]::ConvertFromUtf32($($emoji.Unicode))] = ''" -ForegroundColor White
    Write-Host ""
}
