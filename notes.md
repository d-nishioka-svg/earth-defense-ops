# 地球防衛オペレーション 開発メモ

## アプリ概要
大輝さんの三日坊主をハックする「If-Then：地球防衛オペレーション」アプリ。
Gemini APIを使って毎日新しいキャラクターと壮大な並行世界ミッションを生成し、継続を促す。

## 確定した仕様

### キャラクターライフサイクル
- ミッション（目標＋If-Thenトリガー＋世界観）は一度だけ生成・固定保存
- **毎日アプリを開くたびに新しいキャラクターが登場**
- 同じミッションでも毎日違うキャラが届ける
- 「任務完了」「エネルギー不足（タイムリープ特異点）」どちらもそのキャラが応答
- 10回に1回は「特別通信官」フラグ（間欠強化）

### データ設計
- Googleスプレッドシートで永続化
- **missionsシート**: id, goalText, ifThenTrigger, worldStory, createdDate, status, timing, worldSetting
- **dailyLogsシート**: id, missionId, date, characterName, characterPersonality, result, responseText, doneResponse, skipResponse

### AIキャラクター方向性（確定）
- 「笑わせよう」ではなく「キャラが本気でその世界に生きている」感
- 口調・語尾を最初から最後まで崩さない（DJニャンコのだにゃ語尾、鍛冶頭の職人口調など）
- **実在する日本の固有名詞を必ず1つ入れる**（富士山、甲子園、渋谷、阪神タイガース、吉野家など）
- 固有名詞が臨場感と「ちょっとした面白おかしさ」を生む

### UIテーマ（現在）
- **白基調**：背景 #F3F7FF、カード #FFFFFF
- アクセント: #1A5CCC（深いブルー）
- サブアクセント: #0088BB（シアン）
- スマホ最適化: 16px以上フォント、env(safe-area-inset-bottom)対応

## ファイル構成
- `Code.gs` — GASバックエンド（Gemini APIコール、Spreadsheet操作、doGet）
- `Index.html` — フロントエンドUI（白基調・スマホ最適化）

## GASデプロイ手順
1. clasp push後、GASエディタで「デプロイを管理」→「新しいバージョン」を毎回実施
2. スクリプトプロパティ（1oBPnG8E5...のGAS）：
   - `GEMINI_API_KEY` : Gemini APIキー
   - `SPREADSHEET_ID` : 1GG2q31gGBz6Pk0uLcagCOYMiZy7gienF8_1khGAdXzs

## 科学的アプローチ（組み込まれた心理学）
- **If-Thenプランニング**: 「〇〇したら（If）→△△する（Then）」で行動トリガーを明確化
- **損失回避**: 「世界が崩壊する」という損失フレーミングで継続動機を強化
- **目標の最小化**: サボりそうな時は「1回だけ」という最小ハードルに下げる
- **66日進捗バー**: Phillippaら(2010)「習慣は平均66日で定着」の研究に基づく
- **間欠強化**: 10回に1回の「特別通信官」でスロットマシン効果

## GitHubリポジトリ
https://github.com/d-nishioka-svg/earth-defense-ops.git
