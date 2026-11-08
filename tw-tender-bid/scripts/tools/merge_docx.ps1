# 把投標文件的各份 Word 合併成單一可編輯的 Word 檔,同步輸出 PDF。
# 各部分保留自己的頁首標示,全文連續頁碼(方便評審引述頁次)。需要 Windows 與 Microsoft Word。
# 用法:powershell -ExecutionPolicy Bypass -File merge_docx.ps1 -Manifest <manifest.json>
# manifest 格式見 assets/submission.manifest.example.json(parts 的順序請照招標文件的文件清單)。
param(
    [Parameter(Mandatory = $true)][string]$Manifest
)

$mpath = (Resolve-Path -LiteralPath $Manifest).Path
$m = Get-Content -LiteralPath $mpath -Raw -Encoding UTF8 | ConvertFrom-Json
$BASE = if ($m.base_dir) { [string]$m.base_dir } else { Split-Path $mpath -Parent }
function Resolve-Part([string]$p) { if ([System.IO.Path]::IsPathRooted($p)) { $p } else { Join-Path $BASE $p } }

$OUT_DOCX = Resolve-Part ([string]$m.output_docx)
$OUT_PDF = Resolve-Part ([string]$m.output_pdf)
$PARTS = @($m.parts)
$HEADER_PREFIX = [string]$m.header_prefix
$FOOTER_TEXT = [string]$m.footer_text        # 盲評案請留空,不要放公司名

# 前置檢查:來源檔是否存在、是否被 Word 開著
$blocked = $false
foreach ($p in $PARTS) {
    $path = Resolve-Part ([string]$p.docx)
    if (-not (Test-Path -LiteralPath $path)) { Write-Output ("[缺檔] " + $path); $blocked = $true; continue }
    $lock = Join-Path (Split-Path $path -Parent) ("~`$" + (Split-Path $path -Leaf).Substring(2))
    if (Test-Path -LiteralPath $lock) { Write-Output ("[已開啟,請先關閉] " + $path); $blocked = $true }
}
if ($blocked) { Write-Output ""; Write-Output "請先處理上列檔案後重新執行。"; exit 1 }

Copy-Item -LiteralPath (Resolve-Part ([string]$PARTS[0].docx)) -Destination $OUT_DOCX -Force

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($OUT_DOCX, $false, $false)
    for ($i = 1; $i -lt $PARTS.Count; $i++) {
        $src = Resolve-Part ([string]$PARTS[$i].docx)
        $sel = $word.Selection
        $sel.EndKey(6) | Out-Null          # 6 = wdStory,移到文件結尾
        $sel.InsertBreak(2) | Out-Null     # 2 = wdSectionBreakNextPage
        $sel.InsertFile($src) | Out-Null
    }

    # 頁首頁尾:各節獨立標示,全文連續頁碼
    for ($s = 1; $s -le $doc.Sections.Count; $s++) {
        $sec = $doc.Sections.Item($s)
        $idx = [Math]::Min($s, $PARTS.Count) - 1
        foreach ($h in $sec.Headers) { $h.LinkToPrevious = $false }
        foreach ($f in $sec.Footers) { $f.LinkToPrevious = $false }

        $hdr = $sec.Headers.Item(1).Range
        $hdr.Text = ($HEADER_PREFIX + "  " + [string]$PARTS[$idx].header).Trim()
        $hdr.ParagraphFormat.Alignment = 2      # 靠右
        $hdr.Font.Name = "Calibri"; $hdr.Font.NameFarEast = [string]$m.font_zh
        $hdr.Font.Size = 9; $hdr.Font.Color = 9868950

        $ftr = $sec.Footers.Item(1).Range
        $ftr.Text = $(if ($FOOTER_TEXT) { $FOOTER_TEXT + "  " } else { "" })
        $ftr.ParagraphFormat.Alignment = 1      # 置中
        $ftr.Font.Name = "Calibri"; $ftr.Font.NameFarEast = [string]$m.font_zh
        $ftr.Font.Size = 9; $ftr.Font.Color = 9868950
        $e1 = $sec.Footers.Item(1).Range; $e1.Collapse(0) | Out-Null
        $doc.Fields.Add($e1, 33) | Out-Null     # 33 = wdFieldPage
        $e2 = $sec.Footers.Item(1).Range; $e2.Collapse(0) | Out-Null; $e2.InsertAfter(" / ")
        $e3 = $sec.Footers.Item(1).Range; $e3.Collapse(0) | Out-Null
        $doc.Fields.Add($e3, 26) | Out-Null     # 26 = wdFieldNumPages
        $sec.Footers.Item(1).PageNumbers.RestartNumberingAtSection = $false
    }

    # 更新所有欄位(目錄 PAGEREF、頁次、總頁數)
    $doc.Fields.Update() | Out-Null
    foreach ($sec in $doc.Sections) {
        foreach ($h in $sec.Headers) { $h.Range.Fields.Update() | Out-Null }
        foreach ($f in $sec.Footers) { $f.Range.Fields.Update() | Out-Null }
    }
    $doc.Repaginate()
    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $pdfPath = [string]$OUT_PDF
    $doc.SaveAs([ref]$pdfPath, [ref]17)
    $pages = $doc.ComputeStatistics(2)   # 2 = wdStatisticPages
    $doc.Close([ref]-1)
}
finally {
    $word.Quit()
}
Write-Output ""
Write-Output ("DOCX -> " + (Split-Path $OUT_DOCX -Leaf))
Write-Output ("PDF  -> " + (Split-Path $OUT_PDF -Leaf))
Write-Output ("總頁數:" + $pages + " 頁")
Write-Output "完成。"
