# MMN Takeda 2026 — 講演資料

武田薬品工業 全国研究会 講演「**新ガイドラインを踏まえた多巣性運動ニューロパチーのアップデート**」（35分）の準備資料一式。

## ファイル

| ファイル | 内容 |
|---|---|
| `MMN_update_research_briefing_ja.md` | この1年（2024–2025）のMMNアップデートとEAN/PNS新ガイドライン情報のリサーチ資料。スライド構成案・文献リスト・要確認事項（付録A）付き |
| `MMN_update_2026_skeleton_ja.pptx` | 発表スライドの骨子（25枚・16:9・日本語主体／薬剤・試験名は英語併記） |
| `MMN_update_2026_skeleton_ja_preview.pdf` | 上記pptxのPDFプレビュー |
| `build_deck.py` | pptxを生成するpython-pptxスクリプト（再現用） |

## 主なトピック

- **疫学**：日本全国調査2024（約500名・0.4/10万・M:F 2:1・平均42歳）
- **診断**：EFNS/PNS vs AANEM基準の比較、CBを超えた「拡張基準」(+26%)、神経エコー・NfL
- **新ガイドライン**：EAN/PNS MMNガイドライン（chair S. Goedee、2025–2026見込み・未発表）
- **治療（本命）**：補体阻害薬 empasiprubart（ARDA第2相でIVIg再投与約91%減→EMPASSION第3相）、DNTH103ほか

## 注意

補体阻害薬のデータは学会発表主体（査読論文化は要確認）で、いずれも**国内未承認・臨床試験段階（適応外）**。数値・出典はスライド／リサーチ資料の付録Aで発表前に一次確認を推奨。

## スライド再生成

```bash
pip install python-pptx
python3 build_deck.py   # 出力先パスはスクリプト末尾で調整
```
