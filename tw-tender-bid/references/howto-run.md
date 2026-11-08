# 執行方式(從招標文件到可送件的 PDF)

這份是「動手做」的操作說明。流程、關卡與鐵則見 `SKILL.md`。以下的 `<skill>` 指 skill 資料夾,`<案>` 指你的工作資料夾 `{案名}_投標/`。

## 0. 環境

| 需求 | 說明 |
|---|---|
| Python 3.10 以上 | `pip install -r requirements.txt`(python-docx、matplotlib、pdfplumber、openpyxl、Pillow、pypdf) |
| Windows 加 Microsoft Word | 選用。讀 `.doc`、轉 PDF、更新欄位、合併 Word 都靠它(PowerShell 腳本與 `--finalize`) |
| poppler(`pdftoppm`) | 選用。把 PDF 轉成圖片預覽 |

一律用 `python -I`(隔離模式),避免資料夾裡的同名檔被當成模組載入。處理別人給的招標文件時,下載檔放在獨立的新資料夾,腳本放在另一個資料夾,路徑用參數傳入。

## 1. 設定公司 profile(只做一次)

```
複製 <skill>/assets/company_profile.template.json
填入公司全稱、統編、負責人、地址、電話、信箱、簽章圖檔路徑、公版素材路徑
存成 ~/.claude/tender-profiles/default.json    (或設環境變數 TW_TENDER_PROFILE 指向它)
python -I <skill>/scripts/profile_loader.py    確認讀到的是你的檔案,而不是空白模板
```

簽名與印章圖檔放在 skill 與 repo 之外,profile 只記路徑。

## 2. 收件與拆標

```
mkdir -p <案>/{00_原始,01_拆標,02_大綱,03_草稿,04_審查,05_成品}
把招標文件放進 <案>/00_原始/(原檔不動)
python -I <skill>/scripts/extract_rfp.py <案>/00_原始 --out <案>/01_拆標/text
```

看 `01_拆標/text/_index.json`:出現 `needs_ocr`(掃描檔)或 `no_word_com`(`.doc` 無法轉)的檔案要請對方提供可讀版本。

接著依 `references/rfp-analysis-template.md` 寫 `01_拆標/拆標報告.md`,並用 `assets/format.template.json` 為底寫 `01_拆標/format.json`(格式規格,逐項附須知出處)。

## 3. 挑風格

```
python -I <skill>/scripts/preview_styles.py --list
python -I <skill>/scripts/preview_styles.py --out <案>/03_草稿/style_preview --format <案>/01_拆標/format.json --pdf
```

預覽會套用該案的格式規格。用 `--styles a,b,c` 指定、`--avoid 最近用過的id` 避免與上一案雷同。要新風格:`scripts/make_style.py`。

## 4. 寫 spec、產生 Word

```
複製 <skill>/assets/spec_example.json 為 <案>/03_草稿/spec.json,在 spec 裡填 style、format 與各章內容
python -I <skill>/scripts/build_proposal.py <案>/03_草稿/spec.json --out <案>/03_草稿/服務建議書.docx --finalize
```

- 章節順序照評審項目;缺資料寫 `【待補:項目;負責人】`。
- 金額先用 `scripts/tools/price_utils.py` 換算(含稅、中文大寫),再寫進 spec,不要手打。
- `--finalize` 用 Word 更新欄位,交付的 docx 開檔不會跳更新提示。

## 5. 審查

```
python -I <skill>/scripts/audit_bid.py <案>/03_草稿/服務建議書.docx ^
    --format <案>/01_拆標/format.json --pdf <案>/03_草稿/服務建議書.pdf ^
    --price <最終總價> --project-name "<案名>" --items "評審項目1,評審項目2" ^
    --out <案>/04_審查/audit.json
```

(PowerShell 的續行符號是反引號 `` ` ``,bash 是 `\`。)blocking 必須清零再往下。盲評案加 `--blind`。

## 6. 附件、簽章、合併

1. 機關給的聲明書、切結書、標價單等,優先用原檔填寫。
2. 簽名與印章去背:`python -I <skill>/scripts/tools/prep_seal.py <來源圖> <輸出.png> --ink 26,32,44`(印章通常不加 `--ink`)。**簽章圖檔屬敏感資料,不進公開 repo。** 逐份經使用者確認後,才用 `build_proposal.py --seals` 產出蓋章版。
3. 全部 docx 轉 PDF:`powershell -ExecutionPolicy Bypass -File <skill>/scripts/tools/to_pdf.ps1 -Dir <案>/05_成品/parts`
4. 寫合併清單(範例 `assets/submission.manifest.example.json`),順序照招標文件的文件清單,然後:
   - 單一 PDF 加書籤:`python -I <skill>/scripts/tools/merge_pdf.py <案>/05_成品/manifest.json`
   - 單一可編輯 Word(連續頁碼)並同步輸出 PDF:`powershell -ExecutionPolicy Bypass -File <skill>/scripts/tools/merge_docx.ps1 -Manifest <案>/05_成品/manifest.json`
5. 最後再審查一次合併後的 PDF 頁數與字體,並寫 `05_成品/送件檢查表.md`(份數、簽章、封套、押標金、遞送時間)。

## 7. 常見問題

| 狀況 | 處理 |
|---|---|
| 主控台印出亂碼 | 設 `PYTHONIOENCODING=utf-8`,或在 PowerShell 先 `chcp 65001` |
| 腳本說找不到模組 | 用 `python -I` 時不會自動加入腳本資料夾;直接執行 `scripts/` 下的腳本沒問題,自己寫的程式要 `sys.path.insert(0, "<skill>/scripts")` |
| PowerShell 腳本中文變亂碼 | 檔案要存成 UTF-8 含 BOM(本 skill 的 `.ps1` 已是) |
| 頁首章名出現「錯誤!」 | 已改用 `STYLEREF 1`(不受 Word 介面語言影響);自己加欄位時請同樣用數字寫法 |
| Word 說檔案被鎖住 | 關閉正在開啟的來源檔(資料夾裡會有 `~$` 開頭的暫存檔) |
| 頁數超過上限 | 先看 `format.json` 的 `page_limit_exclude_pages` 是否已扣除封面與目錄;再縮減內容,不要靠縮小字級(會違反規格) |
| 金額大寫 | 用 `price_utils.py`,不要手打 |

## 8. 可重現的示範

repo 的 `examples/demo/run_demo.ps1` 用虛構資料跑完整條自動化流程(產生三份文件、轉 PDF、合併 PDF 與 Word、審查),可以先跑一遍確認環境。
