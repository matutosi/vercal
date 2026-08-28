# vercal プロジェクト

週間の縦型カレンダー (vertical calendar) を作成する Python のコード．
Streamlit の web 版と，手元で実行する版がある．

公開先: <https://vercal.streamlit.app/>

## 主なファイル

- `vercal.py` … カレンダー生成の本体
- `vercal_web.py` … Streamlit の web 版
- `event.py` … 予定 (繰り返しを含む) の扱い
- `schedule.xlsx` … 繰り返しの予定を入力するエクセルの雛形
- `requirements.txt` … 依存パッケージ (streamlit の最小版を指定してある)
- `HackGen35Console-Regular.ttf` … 描画に使うフォント
- `img/` … README 用の画像

## 決めごと

- 設定は web 版の左サイドバーに集約する
  (年，4月始まり/1月始まり，1日の開始・終了時刻，月曜始まり/日曜始まり，左寄せ/右寄せ)．
- **年の既定値は「今の月から決める」** (年度末に翌年を出す)．過去に2回直している箇所なので注意する．
- 繰り返しの予定はエクセルのアップロードで受け取る．書式を変えるときは README の説明も直す．

## 進捗状況

### 現在の状態

- 2026-08-28 11:15
  **同梱フォントのライセンス表記を整えた** (`LICENSE_HackGen.txt` を追加，README に「License ライセンス」節)．
  **README に実際の出力例を追加した** (`img/vercal_output_example.png`)．
  出力例を作る過程で，未報告のバグを2件見つけた (下の「次にやること」)．

- 2026-08-20 08:39
  プロジェクト管理用の `.claude/CLAUDE.md` を新規に設置した．
  実装の最終更新は 2026-03-06 で，年の既定値の算出と streamlit の最小版指定までが入っている．

### 次にやること

- ~~実行例の gif か png を README に追加する~~ **2026-08-28 完了** (`img/vercal_output_example.png`)．
- ~~フォントを同梱しているので，配布時のライセンス表記を確認する~~ **2026-08-28 完了**．
  白源 (HackGen) は **SIL OFL 1.1**，Reserved Font Name は「白源」「HackGen」．
  無改変で再配布しているので，全文 (`LICENSE_HackGen.txt`) を同梱すれば足りる．
  `.gitignore` の `*.ttf` に例外 (`!HackGen35Console-Regular.ttf`) を入れた
  (追跡済みなので消えはしないが，取り違え防止)．
