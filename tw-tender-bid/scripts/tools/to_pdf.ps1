# 將資料夾內所有 .docx 轉為 PDF。轉檔前先更新所有欄位(目錄頁碼、頁次與總頁數、頁首章名、圖表編號)。
# 需要 Windows 與 Microsoft Word。
# 用法:powershell -ExecutionPolicy Bypass -File to_pdf.ps1 -Dir <資料夾>
param(
    [Parameter(Mandatory = $true)][string]$Dir
)

if (-not (Test-Path -LiteralPath $Dir)) { Write-Output "找不到資料夾:$Dir"; exit 1 }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    Get-ChildItem -LiteralPath $Dir -Filter "*.docx" | Where-Object { $_.Name -notlike "~`$*" } | ForEach-Object {
        $src = $_.FullName
        $dst = $src -replace '\.docx$', '.pdf'
        $doc = $word.Documents.Open($src, $false, $false)
        $doc.Fields.Update() | Out-Null
        foreach ($sec in $doc.Sections) {
            foreach ($hf in $sec.Headers) { $hf.Range.Fields.Update() | Out-Null }
            foreach ($hf in $sec.Footers) { $hf.Range.Fields.Update() | Out-Null }
        }
        $doc.Repaginate()
        $doc.Fields.Update() | Out-Null     # 重新分頁後再更新一次,確保頁碼正確
        $doc.SaveAs([ref]$dst, [ref]17)     # 17 = wdFormatPDF
        $doc.Close([ref]-1)                 # -1 = 保留已更新的欄位值
        Write-Output ("PDF -> " + (Split-Path $dst -Leaf))
    }
}
finally {
    $word.Quit()
}
Write-Output "完成。"
