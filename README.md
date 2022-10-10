# MdTestSpec2Excel

## 概要

Markdown形式で記載したテスト仕様書を、Excelに変換するツール。  
MarkdownファイルごとにExcelを出力するのではなく、  
[JXls](https://jxls.sourceforge.net/) 形式のExcelテンプレートに埋め込む形式。

## 実行方法

```bash
md_dir={path_to_markdown_directory}
template_excel={path_to_template_excel_file}
output_excel={path_to_output_excel_file}

./src/converter.main.kts ${md_dir} ${template_excel} ${output_excel}
```

## Docker

T.B.D
