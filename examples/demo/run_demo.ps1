# 用虛構資料跑完整條自動化流程:產生三份文件 -> 轉 PDF -> 合併 PDF -> 合併 Word -> 審查
# 需要 Windows、Python 3.10 以上、Microsoft Word,以及 pip install -r requirements.txt
# 用法:powershell -ExecutionPolicy Bypass -File examples/demo/run_demo.ps1
$ErrorActionPreference = "Stop"
$here = $PSScriptRoot
$skill = Join-Path (Split-Path (Split-Path $here -Parent) -Parent) "tw-tender-bid"
$env:TW_TENDER_PROFILE = Join-Path $here "demo_profile.json"     # 虛構公司資料
$env:PYTHONIOENCODING = "utf-8"
New-Item -ItemType Directory -Force -Path (Join-Path $here "parts"), (Join-Path $here "out") | Out-Null

Write-Output "1) 產生三份範例文件"
foreach ($n in "01_服務建議書", "02_規格需求回應表", "03_報價單") {
    python -I (Join-Path $skill "scripts\build_proposal.py") (Join-Path $skill "assets\spec_example.json") --out (Join-Path $here "parts\$n.docx") | Out-Null
}
Write-Output "2) 轉 PDF"
powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $skill "scripts\tools\to_pdf.ps1") -Dir (Join-Path $here "parts")
Write-Output "3) 合併成單一 PDF(含書籤)"
python -I (Join-Path $skill "scripts\tools\merge_pdf.py") (Join-Path $here "manifest.json")
Write-Output "4) 合併成單一可編輯 Word,並輸出 PDF"
powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $skill "scripts\tools\merge_docx.ps1") -Manifest (Join-Path $here "manifest.json")
Write-Output "5) 審查(示範檔刻意留有【待補】,會被抓出來)"
python -I (Join-Path $skill "scripts\audit_bid.py") (Join-Path $here "parts\01_服務建議書.docx") --price 1062600
Write-Output "完成。輸出在 examples/demo/out/"
