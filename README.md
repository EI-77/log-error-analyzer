# Log Error Analyzer

複数のログファイルを読み込み，`WARNING`，`ERROR`，`CRITICAL` などの重要ログを抽出し，ログレベル別・カテゴリ別・ファイル別・日付別に集計したExcelレポートを出力するPythonツールです．

`input` フォルダに配置した `.log` / `.txt` ファイルをまとめて読み込み，ログ行から日時，ログレベル，メッセージを抽出します．
抽出した重要ログは，メッセージ内のキーワードをもとに `database`，`timeout`，`permission`，`api` などのカテゴリへ分類し，`log_analysis_report.xlsx` に出力します．

通常ログである `INFO` は集計対象から除外し，形式が壊れているログ行は `parse_errors` シートに分離して出力します．

このツールは，大量のログファイルを目視で確認する作業を減らし，障害調査や運用確認に必要な情報をExcelで確認しやすくするためのツールです．

---

## 目次

- [主な機能](#主な機能)
- [使用技術](#使用技術)
- [想定用途](#想定用途)
- [ディレクトリ構成](#ディレクトリ構成)
- [実行方法](#実行方法)
- [初回実行について](#初回実行について)
- [入力ログの形式](#入力ログの形式)
- [サンプルログファイル](#サンプルログファイル)
- [設定項目](#設定項目)
- [出力ファイル](#出力ファイル)
- [ログ分類の考え方](#ログ分類の考え方)
- [severity について](#severity-について)
- [文字コードについて](#文字コードについて)
- [注意事項](#注意事項)
- [活用方法](#活用方法)
- [Requirements](#requirements)
- [English](#english)

---

## 主な機能

* `.log` ファイルの読み込み
* `.txt` ファイルの読み込み
* 複数ログファイルの一括解析
* `WARNING`，`ERROR`，`CRITICAL` の抽出
* `INFO` などの通常ログの除外
* 正規表現によるログ行解析
* 日時，ログレベル，メッセージの抽出
* ログメッセージのカテゴリ分類
* ログレベル別集計
* カテゴリ別集計
* ファイル別集計
* 日付別集計
* 解析できなかったログ行の分離
* 処理結果サマリーのExcel出力
* 複数シート構成のExcelレポート出力
* 初回実行用サンプルログの自動作成
* `config.json` による解析条件の変更
* UTF-8 / UTF-8 BOM / CP932 の文字コード対応
* Excelシートの列幅調整
* Excelシートのヘッダー装飾
* Excelシートのオートフィルター設定

---

## 使用技術

* Python
* pandas
* XlsxWriter

---

## 想定用途

このツールは，以下のような作業を自動化するためのものです．

* アプリケーションログからエラーを抽出する
* サーバーログから警告や障害ログを抽出する
* バッチ処理ログの失敗内容を確認する
* 複数ログファイルをまとめて確認する
* 大量ログから `ERROR` や `CRITICAL` だけを抜き出す
* エラー内容をカテゴリ別に整理する
* どのファイルでエラーが多いか確認する
* どの日にエラーが多いか確認する
* 解析できないログ行を確認する
* 障害調査用のExcelレポートを作成する
* ログ確認作業を自動化する
* 運用確認作業の補助ツールとして使う
* ログ監視・通知ツールのベースとして使う

---

## ディレクトリ構成

```text
log-error-analyzer/
├─ log_error_analyzer.py
├─ config.json
├─ requirements.txt
├─ README.md
├─ input/
│  ├─ sample_app.log
│  ├─ sample_server.log
│  └─ sample_legacy.txt
└─ output/
   └─ log_analysis_report.xlsx
```

### フォルダの役割

| フォルダ     | 内容                   |
| -------- | -------------------- |
| `input`  | 解析対象のログファイルを配置するフォルダ |
| `output` | 解析レポートの出力先           |

---

## 実行方法

### 1. リポジトリを取得

```bash
git clone <repository_url>
```

### 2. ディレクトリへ移動

```bash
cd log-error-analyzer
```

### 3. 必要ライブラリをインストール

```bash
pip install -r requirements.txt
```

### 4. ログファイルを配置

解析したいログファイルを `input` フォルダに配置します．

対応形式は以下です．

```text
.log
.txt
```

例：

```text
input/app.log
input/server.log
input/batch_log.txt
```

### 5. config.json を確認

`config.json` で，解析対象の拡張子，抽出対象ログレベル，日時形式，出力ファイル名，カテゴリ分類用キーワードを設定します．

### 6. プログラムを実行

```bash
python log_error_analyzer.py
```

または

```bash
py log_error_analyzer.py
```

### 7. 出力ファイルを確認

実行後，`output` フォルダに以下のExcelレポートが出力されます．

```text
output/
└─ log_analysis_report.xlsx
```

---

## 初回実行について

初回実行時，以下のファイルやフォルダが存在しない場合は自動作成されます．

```text
config.json
input/sample_app.log
input/sample_server.log
input/sample_legacy.txt
output/
```

そのため，最初は自分でログファイルを用意しなくても，以下のコマンドだけで動作確認できます．

```bash
python log_error_analyzer.py
```

サンプルログには，通常ログ，警告ログ，エラーログ，重大ログ，形式が壊れたログ行が含まれています．

これにより，重要ログの抽出，カテゴリ分類，集計，parse error の分離，Excelレポート出力まで一通り確認できます．

---

## 入力ログの形式

このツールは，初期設定では以下の形式のログを解析します．

```text
YYYY-MM-DD HH:MM:SS LEVEL message
```

例：

```text
2026-06-15 09:03:45 WARNING Disk usage is high: 82%
2026-06-15 09:05:10 ERROR Failed to connect database
2026-06-15 09:30:00 CRITICAL Out of memory while processing batch job
```

この形式から，以下の情報を抽出します．

| 項目       | 例                            |
| -------- | ---------------------------- |
| datetime | `2026-06-15 09:05:10`        |
| level    | `ERROR`                      |
| message  | `Failed to connect database` |

---

## サンプルログファイル

### sample_app.log

アプリケーションログを想定したサンプルです．

含まれる内容の例：

* アプリケーション起動
* 設定ファイル読み込み
* ディスク使用量警告
* データベース接続失敗
* SQLエラー
* API応答遅延
* 権限エラー
* メモリエラー
* ユーザー情報未検出
* 形式が壊れたログ行

---

### sample_server.log

サーバー・APIログを想定したサンプルです．

含まれる内容の例：

* サーバー起動
* HTTPリクエスト遅延
* APIリクエスト失敗
* upstream server 接続拒否
* ディスク空き容量低下
* ファイル未検出
* データベース障害
* 外部API timeout
* 形式が壊れたログ行

---

### sample_legacy.txt

古いシステムやバッチ処理などで出力される `.txt` 形式ログを想定したサンプルです．

このファイルは，`.log` 以外の拡張子でもログ解析できることを確認するために用意しています．

含まれる内容の例：

* 夜間バッチ開始
* メモリ使用量警告
* 出力ファイル書き込み失敗
* ディスク容量不足
* 夜間バッチ終了


---

## 設定項目
動作設定は `config.json` で変更できます．

| 項目                  | 内容                           |
| ------------------- | ---------------------------- |
| `target_extensions` | 解析対象とするファイル拡張子               |
| `target_levels`     | 抽出対象とするログレベル                 |
| `datetime_format`   | ログ日時の解析形式                    |
| `output_file`       | `output` フォルダに出力するExcelファイル名 |
| `category_keywords` | ログメッセージをカテゴリ分類するためのキーワード設定   |

### サンプル設定

```json
{
  "target_extensions": [
    ".log",
    ".txt"
  ],
  "target_levels": [
    "WARNING",
    "ERROR",
    "CRITICAL"
  ],
  "datetime_format": "%Y-%m-%d %H:%M:%S",
  "output_file": "log_analysis_report.xlsx",
  "category_keywords": {
    "database": [
      "database",
      "db",
      "sql",
      "mysql",
      "postgres"
    ],
    "timeout": [
      "timeout",
      "timed out"
    ],
    "permission": [
      "permission",
      "denied",
      "unauthorized",
      "forbidden"
    ],
    "not_found": [
      "not found",
      "missing",
      "404"
    ],
    "memory": [
      "memory",
      "out of memory",
      "oom"
    ],
    "connection": [
      "connection",
      "connect",
      "network",
      "socket"
    ],
    "api": [
      "api",
      "http",
      "request",
      "response"
    ],
    "disk": [
      "disk",
      "storage",
      "no space"
    ]
  }
}
```

---

## target_extensions について

`target_extensions` は，解析対象にするファイル拡張子を指定する設定です．

標準設定では，以下のファイルを解析します．

```text
.log
.txt
```

例：

```json
{
  "target_extensions": [
    ".log",
    ".txt"
  ]
}
```

この設定の場合，`input` フォルダ内の `.log` と `.txt` ファイルが解析対象になります．

---

## target_levels について

`target_levels` は，抽出対象にするログレベルを指定する設定です．

標準設定では，以下のログレベルを抽出します．

```text
WARNING
ERROR
CRITICAL
```

例：

```json
{
  "target_levels": [
    "WARNING",
    "ERROR",
    "CRITICAL"
  ]
}
```

`INFO` や `DEBUG` は，ログ形式としては正しくても抽出対象外として無視されます．
parse error にはなりません．

---

## datetime_format について

`datetime_format` は，ログ行の日時を解析するための形式です．

標準設定では以下の形式です．

```text
%Y-%m-%d %H:%M:%S
```

対応するログ例：

```text
2026-06-15 09:05:10 ERROR Failed to connect database
```

日時形式が設定と合わない場合，その行は `parse_errors` シートに出力されます．

---

## category_keywords について

`category_keywords` は，ログメッセージをカテゴリ分類するための設定です．

例：

```json
{
  "database": [
    "database",
    "db",
    "sql"
  ],
  "timeout": [
    "timeout",
    "timed out"
  ]
}
```

たとえば，以下のログメッセージは `database` に分類されます．

```text
Failed to connect database
```

以下のログメッセージは `timeout` に分類されます．

```text
Timeout while calling external API
```

どのカテゴリにも一致しない場合は，`other` に分類されます．

複数カテゴリに一致する場合は，`config.json` で先に定義されているカテゴリが採用されます．

---

## 出力ファイル

### log_analysis_report.xlsx

ログ解析結果をまとめて出力するExcelファイルです．

例：

```text
output/log_analysis_report.xlsx
```

このファイルには，以下のシートが含まれます．

| シート名               | 内容           |
| ------------------ | ------------ |
| `summary`          | 処理結果の概要      |
| `log_details`      | 抽出された重要ログの詳細 |
| `level_summary`    | ログレベル別件数     |
| `category_summary` | カテゴリ別件数      |
| `file_summary`     | ファイル別件数      |
| `timeline_summary` | 日付別件数        |
| `parse_errors`     | 解析できなかったログ行  |
| `config`           | 解析時に使用した設定   |

---

### summary シート

処理結果の概要を出力するシートです．

出力例：

| item              | value |
| ----------------- | ----: |
| checked_files     |     3 |
| checked_lines     |    30 |
| target_logs       |    19 |
| warning_count     |     6 |
| error_count       |    11 |
| critical_count    |     2 |
| parse_error_count |     2 |

`started_at` と `finished_at` も出力されます．

---

### log_details シート

抽出された重要ログを出力するシートです．

出力列：

| 列名            | 内容              |
| ------------- | --------------- |
| `file_name`   | 元ログファイル名        |
| `line_number` | 元ログファイル上の行番号    |
| `datetime`    | ログ発生日時          |
| `date`        | ログ発生日           |
| `level`       | ログレベル           |
| `severity`    | 重要度             |
| `category`    | メッセージから分類したカテゴリ |
| `message`     | ログメッセージ         |
| `raw_line`    | 元のログ行           |

出力例：

| file_name         | line_number | datetime            | level    | severity | category | message                                  |
| ----------------- | ----------: | ------------------- | -------- | -------- | -------- | ---------------------------------------- |
| sample_app.log    |           4 | 2026-06-15 09:05:10 | ERROR    | high     | database | Failed to connect database               |
| sample_app.log    |          10 | 2026-06-15 09:30:00 | CRITICAL | high     | memory   | Out of memory while processing batch job |
| sample_server.log |          10 | 2026-06-15 10:15:32 | ERROR    | high     | timeout  | Timeout while calling external API       |

`file_name` と `line_number` を見れば，元ログファイルのどの行を確認すればよいか分かります．

---

### level_summary シート

ログレベル別の件数を出力するシートです．

出力例：

| level    | count |
| -------- | ----: |
| CRITICAL |     2 |
| ERROR    |    11 |
| WARNING  |     6 |

---

### category_summary シート

カテゴリ別の件数を出力するシートです．

出力列：

| 列名               | 内容         |
| ---------------- | ---------- |
| `category`       | エラーカテゴリ    |
| `count`          | カテゴリ全体の件数  |
| `warning_count`  | WARNING件数  |
| `error_count`    | ERROR件数    |
| `critical_count` | CRITICAL件数 |

出力例：

| category   | count | warning_count | error_count | critical_count |
| ---------- | ----: | ------------: | ----------: | -------------: |
| database   |     4 |             0 |           3 |              1 |
| api        |     3 |             2 |           1 |              0 |
| disk       |     3 |             2 |           1 |              0 |
| connection |     2 |             1 |           1 |              0 |

---

### file_summary シート

ファイル別の解析結果を出力するシートです．

出力例：

| file_name         | checked_lines | target_logs | warning_count | error_count | critical_count | parse_error_count |
| ----------------- | ------------: | ----------: | ------------: | ----------: | -------------: | ----------------: |
| sample_app.log    |            15 |           9 |             3 |           5 |              1 |                 1 |
| sample_legacy.txt |             5 |           3 |             1 |           2 |              0 |                 0 |
| sample_server.log |            10 |           7 |             2 |           4 |              1 |                 1 |

このシートにより，どのログファイルに重要ログや parse error が多いか確認できます．

---

### timeline_summary シート

日付別の件数を出力するシートです．

出力例：

| date       | warning_count | error_count | critical_count | total_count |
| ---------- | ------------: | ----------: | -------------: | ----------: |
| 2026-06-14 |             1 |           2 |              0 |           3 |
| 2026-06-15 |             5 |           9 |              2 |          16 |

---

### parse_errors シート

想定ログ形式に一致しなかった行を出力するシートです．

出力列：

| 列名              | 内容           |
| --------------- | ------------ |
| `file_name`     | 元ログファイル名     |
| `line_number`   | 元ログファイル上の行番号 |
| `raw_line`      | 解析できなかった元の行  |
| `error_message` | 解析できなかった理由   |

出力例：

| file_name         | line_number | raw_line                                 | error_message                           |
| ----------------- | ----------: | ---------------------------------------- | --------------------------------------- |
| sample_app.log    |          13 | BROKEN LOG LINE WITHOUT EXPECTED FORMAT  | Line does not match expected log format |
| sample_server.log |           9 | INVALID LINE: missing datetime and level | Line does not match expected log format |

`INFO` のような通常ログは parse error ではありません．
ログ形式が正しく，単に抽出対象ではないだけの行は無視されます．

---

### config シート

解析時に使用した設定を出力するシートです．

出力例：

| key               | value                    |
| ----------------- | ------------------------ |
| target_extensions | .log, .txt               |
| target_levels     | WARNING, ERROR, CRITICAL |
| datetime_format   | %Y-%m-%d %H:%M:%S        |
| output_file       | log_analysis_report.xlsx |

このシートにより，後からExcelレポートを見たときに，どの条件で解析した結果なのかを確認できます．

---

## ログ分類の考え方

このツールでは，ログ行を以下のように分類します．

```text
空行
        → 無視

ログ形式に一致しない行
        → parse_errors シート

ログ形式に一致するが target_levels ではない行
        → 無視

ログ形式に一致し，target_levels に含まれる行
        → log_details シート
        → 各種summaryの集計対象
```

例：

```text
2026-06-15 09:00:01 INFO Application started
```

この行はログ形式としては正しいですが，標準設定では `INFO` が抽出対象ではないため無視されます．

```text
BROKEN LOG LINE WITHOUT EXPECTED FORMAT
```

この行は想定フォーマットに一致しないため，`parse_errors` シートに出力されます．

---

## severity について

`severity` は，ログレベルから付与する重要度です．

標準設定では以下のように分類します．

| level      | severity |
| ---------- | -------- |
| `CRITICAL` | `high`   |
| `ERROR`    | `high`   |
| `WARNING`  | `medium` |

---

## 文字コードについて

ログファイル読み込み時は，以下の文字コードを順番に試します．

```text
utf-8-sig
utf-8
cp932
```

これにより，日本語環境のログファイルでも読み込める可能性が高くなります．

---

## 注意事項

* 解析対象ファイルは `input` フォルダ直下に配置してください．
* 初期版ではサブフォルダ内の再帰探索には対応していません．
* 標準では `.log` と `.txt` を解析対象にします．
* 対象拡張子は `config.json` の `target_extensions` で変更できます．
* 標準では `WARNING`，`ERROR`，`CRITICAL` を抽出します．
* 抽出対象ログレベルは `config.json` の `target_levels` で変更できます．
* ログ形式は `YYYY-MM-DD HH:MM:SS LEVEL message` を想定しています．
* 想定形式に一致しない行は `parse_errors` シートに出力されます．
* `INFO` や `DEBUG` は parse error ではなく，抽出対象外として無視されます．
* カテゴリ分類はキーワードベースです．
* 複数カテゴリに一致した場合は，先に定義されているカテゴリが採用されます．
* どのカテゴリにも一致しない場合は `other` に分類されます．
* `output` フォルダ内の生成ファイルは実行結果なので，Git管理対象から外しても問題ありません．
* `config.json` やサンプルファイルが既に存在する場合，自動作成処理では上書きされません．
* 同名の出力ファイルが既にある場合は，実行時に上書きされます．
* `.gz` 圧縮ログには未対応です．
* リアルタイム監視には未対応です．
* Slack通知やメール通知には未対応です．
* GUIには未対応です．

---

## 活用方法

* 障害発生後のログ確認
* サーバーログのエラー抽出
* アプリケーションログの調査
* バッチ処理ログの失敗確認
* `ERROR` / `WARNING` の件数集計
* 複数ログファイルの横断確認
* エラー原因カテゴリの整理
* 日別のエラー傾向確認
* 運用確認作業の自動化
* 障害調査用Excelレポート作成
* ログ監視・通知ツールへの拡張ベース

---

## Requirements

```text
pandas
xlsxwriter
```

---

# English

## Overview

Log Error Analyzer is a Python tool that reads multiple log files, extracts important log lines such as `WARNING`, `ERROR`, and `CRITICAL`, classifies them by category, and exports a multi-sheet Excel report.

The tool reads `.log` and `.txt` files from the `input` folder.
It parses each log line, extracts datetime, level, and message, and summarizes the results by log level, category, file, and date.

Malformed log lines are exported to a separate `parse_errors` sheet.

---

## Features

* Read `.log` files
* Read `.txt` files
* Analyze multiple log files at once
* Extract `WARNING`, `ERROR`, and `CRITICAL` logs
* Ignore normal `INFO` logs
* Parse log lines with a regular expression
* Extract datetime, log level, and message
* Classify log messages by keyword
* Create level summary
* Create category summary
* Create file summary
* Create timeline summary
* Separate parse errors
* Export a multi-sheet Excel report
* Create sample files on first run
* Support common text encodings
* Apply basic Excel formatting

---

## Directory Structure

```text
log-error-analyzer/
├─ log_error_analyzer.py
├─ config.json
├─ requirements.txt
├─ README.md
├─ input/
│  ├─ sample_app.log
│  ├─ sample_server.log
│  └─ sample_legacy.txt
└─ output/
   └─ log_analysis_report.xlsx
```

---

## How to Use

### 1. Install requirements

```bash
pip install -r requirements.txt
```

### 2. Place log files

Place `.log` or `.txt` files in the `input` folder.

```text
input/app.log
input/server.log
input/batch_log.txt
```

### 3. Edit config.json

Set target extensions, target log levels, datetime format, output file name, and category keywords.

### 4. Run the script

```bash
python log_error_analyzer.py
```

---

## Output

Generated files are saved in the `output` folder.

```text
output/log_analysis_report.xlsx
```

The Excel report contains the following sheets.

| Sheet              | Description                     |
| ------------------ | ------------------------------- |
| `summary`          | Summary of the analysis result  |
| `log_details`      | Extracted important log records |
| `level_summary`    | Count by log level              |
| `category_summary` | Count by error category         |
| `file_summary`     | Count by source file            |
| `timeline_summary` | Count by date                   |
| `parse_errors`     | Lines that could not be parsed  |
| `config`           | Configuration used for analysis |

---

## Supported Log Format

This version supports the following log format.

```text
YYYY-MM-DD HH:MM:SS LEVEL message
```

Example:

```text
2026-06-15 09:05:10 ERROR Failed to connect database
```

This line is parsed as follows.

```text
datetime: 2026-06-15 09:05:10
level: ERROR
message: Failed to connect database
```

---

## Result Classification

The tool classifies log lines as follows.

```text
Empty lines
        -> ignored

Lines that do not match the expected format
        -> parse_errors

Lines that match the format but are not target levels
        -> ignored

Lines that match the format and are target levels
        -> log_details
        -> included in summaries
```

---

## Limitations

* This version reads files directly under the `input` folder.
* Recursive directory scanning is not supported.
* `.gz` compressed logs are not supported.
* Automatic log format detection is not supported.
* Apache, nginx, syslog, and JSON log formats are not supported.
* Multi-line stack traces are not supported.
* Real-time monitoring is not supported.

---

## Use Cases

* Extract errors from application logs
* Analyze server logs
* Review batch job failures
* Summarize `WARNING`, `ERROR`, and `CRITICAL` records
* Create Excel reports for incident investigation
* Reduce manual log checking work
* Use as a base for future log monitoring tools
