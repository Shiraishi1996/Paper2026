# PaperForge Deterministic Web — LLM 0%

ブラウザで動く、生成AI・LLM完全不使用の研究素材整理／論文構築Webアプリです。

## 対応素材
- PDF（テキストPDF／任意OCRでスキャンPDF）
- Word (.docx)（本文・表・埋め込み画像）
- CSV / TSV / Excel
- TXT / Markdown
- PNG / JPEG / TIFF / WebP

## 主要機能
- 見出しルールによる Abstract / Introduction / Methods / Results / Discussion / Conclusion の抽出
- Word内画像の抽出
- 日本語＋英語Tesseract OCR（任意）
- 記述統計・Spearman相関
- TF-IDFによる図の配置候補判定
- DOI抽出
- Crossref DOI実在照合（任意、LLMではない）
- 投稿先scopeのTF-IDF類似度
- Word / PDF / LaTeX / Markdown出力
- ファイルSHA-256とページ／段落位置を監査ログとして保持

## ローカル起動
```bash
pip install -r requirements.txt
streamlit run app.py
```

OCRを使う場合はOS側にTesseractとPopplerが必要です。

## Streamlit Community Cloudで公開
1. このフォルダの `app.py`, `requirements.txt`, `packages.txt`, `.streamlit/config.toml` をGitHubリポジトリへ置く。
2. Streamlit Community Cloudで `Create app` を選ぶ。
3. Repositoryを指定し、Main file pathを `app.py` にする。
4. Deployする。

APIキーやLLM用Secretは不要です。

## Dockerで公開する場合
```bash
docker build -t paperforge .
docker run -p 8501:8501 paperforge
```
