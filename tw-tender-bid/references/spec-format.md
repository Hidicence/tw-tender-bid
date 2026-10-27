# build_proposal.py 的 spec 格式

```
python -I scripts/build_proposal.py spec.json --out 05_成品/服務建議書.docx [--seals]
```

`--seals` 才會插入負責人簽名與印章。**未經使用者逐份確認不要加。** 預設簽章區留空白底線。範例:`assets/spec_example.json`。

## 頂層欄位

| 欄位 | 型別 | 說明 |
|---|---|---|
| `doc_kind` | 字串 | 封面副標,預設「服務建議書」(可改「企劃書」「計畫書」) |
| `style` | 字串 | 風格包 id(`preview_styles.py --list` 列出;`navy`、`green` 為舊別名) |
| `format` | 路徑或物件 | 招標文件的格式規格(`format.json`),規格優先於風格;含 `numbering` 時章節標號自動產生,spec 內的章標題不要手打標號 |
| `toc` | 布林 | 是否產生目錄(章與節自動建,點線引導加頁碼欄位) |
| `toc_levels` | 數字 | 目錄層級,1 只列章,2 列章與節(預設 2) |
| `cover` | 布林 | 是否產生封面 |
| `case` | 物件 | 案件資訊,見下表 |
| `sections` | 陣列 | 內容區塊,依序輸出 |

### case

| 欄位 | 說明 |
|---|---|
| `title` | 案名全稱(封面主標) |
| `short` | 簡稱(頁首) |
| `client` | 機關名稱 |
| `subtitle` | 案號或副標 |
| `doc_sub` | 封面補充說明 |
| `cover_date` | 封面日期,例:`中華民國 115 年 11 月` |
| `sign_date` | 簽章區日期,留空則為空白欄位 |
| `blind` | `true` 時封面、簽章區與頁尾不顯示公司名、統編、負責人(盲評案) |
| `cover_image` | 封面主視覺圖檔路徑(選用,例如 AI 生成的主視覺);不存在則略過 |
| `partner` | 協力單位,例:`{"name": "協力單位全名"}`,封面會顯示;沒有就 `null` |

## section 類型

| type | 欄位 | 說明 |
|---|---|---|
| `banner` | `text`、`sub`、`new_page` | 章大標(深靛藍底),也是目錄來源 |
| `subh` | `text`、`num`、`new_page` | 次標題 |
| `subsubh` | `text` | 三級標題 |
| `body` | `text` | 內文段落 |
| `bullets` | `items` | 項目符號清單 |
| `numlist` | `items`、`start` | 編號清單 |
| `table` | `headers`、`rows`、`widths`(cm)、`size`、`center_from`、`caption` | 表格。首欄以「一、」開頭視為分類列,含「小計」「合計」「總計」「稅金」的列自動套色 |
| `kv` | `pairs`(`[["標籤","值"],...]`) | 兩欄資料表 |
| `hbox` | `text` | 重點框 |
| `quote` | `title`、`text` | 引用框 |
| `image` | `path`、`width`(cm)、`caption` | 圖片;檔案不存在會輸出【待補】提醒 |
| `chart` | `kind`(`pie`/`bar`/`gantt`)、`data`、`caption`、`unit`、`width`;`caption` 不用寫「圖 N」,會自動編號 | 圖表。pie/bar 的 data 是 `[{"label","value"}]`;gantt 是 `[{"task","start","end"}]` |
| `lead` | `text` | 導言:比內文大一級、主色,放在章標題後的一兩句重點 |
| `stats` | `items`(`[{"value","unit","label"}]`,最多 4 個) | 重點數字卡(年資、案件數、團隊人數等,只放有證據的數字) |
| `callout` | `title`、`text`、`kind`(`note`、`key`、`warn`) | 提示框,左側色條,不用圖示 |
| `steps` | `items`(`[{"title","text"}]`) | 步驟流程,左側大號序號 |
| `twocol` | `left`、`right`(各為 `{"title","items"}`) | 雙欄對照(現況與做法、需求與回應) |
| `divider` | `number`、`title`、`subtitle` | 章節扉頁(整頁),本身是章標題,會進目錄與導覽窗格;用它就不要再放同章的 `banner` |
| `page_break` | | 分頁 |
| `sign_block` | `date` | 投標廠商簽章區(公司資料來自公司 profile,見 SKILL.md「公司 profile」) |

## 輸出與更新欄位

`python -I scripts/build_proposal.py spec.json --out 成品.docx --finalize` 會用本機 Word 更新所有欄位(目錄頁碼、頁碼、頁首章名、圖表編號)並存檔,交付的 docx 開檔時就不會跳出「是否更新欄位」的提示。沒有 Word 時略過,開檔時 Word 會詢問是否更新。

## 慣例

- 章節順序依評審項目,banner 的 `sub` 欄放對應的評審項目名稱,審查時好對照。
- 設了 `format.numbering` 之後,banner、subh、subsubh 的標號由程式依規定產生(會先移除你手打的標號),所以 `subh` 的 `num` 欄不再使用。
- 缺資料寫 `【待補:項目;負責人】`,`audit_bid.py` 會抓出來;交件前必須清零。
- 價格表的數字從單一來源產生(例如先在 Python 裡算好再寫入 spec),不要在各章手打。
- 需要更細的版面(分鏡腳本表、組織圖、特殊表格),直接 `import bid_kit as K`,寫自己的 `gen_*.py`。
