# File Masking Tool — Handoff

更新日: 2026-10-04

## 目的と実行環境

本番ファイル内の個人情報を、組織の管理下でローカルにマスキングするPowerShellツールです。固定長テキスト、CSV、Excel `.xls` / `.xlsx` を対象にし、Pythonや外部サービスは使いません。レイアウトに指定した項目だけを変換し、FILLERや未指定項目は元の値を保持します。未知の変換値や不正なレコードは成功出力に通しません。

前提はWindows 11、Windows PowerShell 5.1、`.NET Framework` の `Add-Type` C#実行時コンパイルです。Excelを処理する場合は登録済みOffice Excel COMも必要です。

## 現在の実装と検証

- 固定長とCSVのレコード処理は、隣接する `src\Masking\MaskingCore.cs` を実行時にメモリ内コンパイルする方式です。PowerShellのCLI、設定、変換表読込、Runner、Report、Excel COMは維持しています。
- コンパイル失敗は通常実行開始前に明示し、終了コード2で失敗します。PowerShell実装への黙示フォールバックはありません。`-Help` と `-ValidateOnly` はコアのコンパイルを行いません。
- 固定長テキストとCSVのファイル文字コードはcp932（コードページ932）に限定し、それ以外は設定検証時に拒否します。変換表の文字コードはファイル入力とは別設定です。
- 固定長ではFILLERのバイト保持、厳密なcp932処理、終端検証、分割フィールドの一括変換、決定的サンプリングを実装しています。
- CSVではストリーミング解析、引用状態・埋め込み改行・エスケープ引用符・BOMの保持、固定幅NUL埋め、決定的サンプリングを実装しています。
- 最終テスト結果: **59 passed、0 failed**。Excel COM統合テストは現在のPCにCOM登録がなくBLOCKEDです（成功には数えていません）。
- cp932限定後の300,000件合成測定: C#コンパイル0.416秒、合成データ生成4.405秒、変換表読込38.261秒、固定長0.416秒（721,153.85件/秒）、CSV 0.351秒（854,700.85件/秒）。Excel測定はCOMがなくBLOCKEDです。これらはこのPCでの測定値であり、SLAではありません。

## 別PCへ持っていくもの

**最も確実なのは、`C:\work\NS\FileMaskingTool` フォルダー全体をコピーすることです。** Gitのブランチと未コミット変更も引き継ぐ場合は、隠し項目を含めて `.git` と `.superpowers` を省かずコピーしてください。特定のファイルだけを選んでコピーすると、未コミット変更や作業計画を落とす可能性があります。

特に確認するもの:

- `MaskingTool.ps1`
- `src\Masking\` 全体（`MaskingCore.cs` と `CompiledCore.psm1` を含む）
- `tests\` 全体
- `README.md`、`HANDOFF.md`、`config.example.json`
- `docs\superpowers\` の設計・計画書
- `.superpowers\sdd\2026-10-04-compiled-masking-core\` の進捗台帳・タスクブリーフ
- `.git` と、存在する場合は `.gitignore` 等の隠しファイル

本番入力ファイルや個人情報を含む変換表は、組織の承認なしに別PCへコピーしないでください。ツールの開発継続にそれらは不要です。例示設定とテストは合成データを使います。

## 別PCでの再開

1. Windows 11とWindows PowerShell 5.1が使えることを確認します。Excel統合も行う場合はOffice Excel COM登録を確認します。
2. コピーしたプロジェクトのルートで状態を確認します。

   ```powershell
   git status --short --branch
   ```

   作業ブランチは `feature/file-masking-tool` です。複製時には未コミット変更が多数残っています。既存変更を破棄・リセットしないでください。

3. C#実行時コンパイルと機能を確認します。

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1
   ```

4. 30万件の合成測定を再実行する場合:

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Measure-Performance.ps1
   ```

5. 実データを扱う前に、設定を `-ValidateOnly` で検証し、組織のデータ持ち出しルールと変換表の管理方法を確認します。

## 未完了・注意点

- 独立コードレビューで指摘された固定長出力衝突時の既存ファイル保護、CSVサンプルのデコード例外サニタイズ、および文字コード前提の問題に対応しました。回帰テストを含む全テストを再実行済みです。
- Git author identityが未設定のため、作業コミットは保留中です。Gitの設定を勝手に変更していません。コミットする場合は、利用者が承認する氏名・メールアドレスをローカル設定してください。
- Excel COM統合はOffice Excel COMが登録されたPCで別途実行してください。
- 本番ファイルやマッピングファイルを生成AIへ渡さず、未指定領域を含めた出力全体を利用前に確認してください。

設計の根拠は `docs\superpowers\specs\2026-10-04-compiled-masking-core-design.md`、実装計画は `docs\superpowers\plans\2026-10-04-compiled-masking-core.md` です。
