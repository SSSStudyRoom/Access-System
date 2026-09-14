# SSS Education 自習室管理システム 引継ぎ資料

最終更新：2026-09-14（コードを実際に全文精読した上で作成。以前存在した同名ファイルは失われていたため新規作成）

このドキュメントは、リポジトリ内の各ファイルを実際に読み込んで確認した「現状の仕様」をまとめたものです。
開発者が複数世代にわたっているため、意図が完全には分からない箇所もそのまま正直に記載しています。

---

## 1. システム全体像

塾（SSS Education）の自習室運営システム。3つの実行環境にまたがる。

```
[受付iPad] index.html ──POST(scan)──┐
                                     │
[生徒スマホ] dashboard.html ──GET/POST─┤→ GAS Web App (コード.gs / AbsenceCheck.gs)
                                     │      ├─ スプレッドシート（生徒マスタ・ログ・投稿等）
[スプレッドシートUI] draftDialog.html ┘      ├─ Googleカレンダー（生徒ごとの個別カレンダー）
   (google.script.run で直接呼び出し)         └─ Gmail / Google Chat 通知
```

- **index.html / dashboard.html** は GitHub Pages で静的ホスティングされている（`https://ahmadtanzeel.github.io/-dev-webpages-/` 配下、コード.gs 内の `STUDENT_DASHBOARD_BASE_URL` 参照）。
- **コード.gs / AbsenceCheck.gs / draftDialog.html / appsscript.json** は同一の Google Apps Script プロジェクトにバインドされ、スプレッドシートから直接デプロイされる。GAS は同一プロジェクト内なら `.gs` ファイルをまたいでもグローバルスコープが共有されるため、`コード.gs` の定数・関数を `AbsenceCheck.gs` からそのまま参照できる（`IDX_PROFILE`, `CALENDAR_CONFIG`, `ensureSheet` など）。
- フロントエンドとGASの通信は `fetch` ベース。POSTは CORS プリフリcourt（OPTIONS）を避けるため、JSONを送るにもかかわらず意図的に `Content-Type: text/plain;charset=utf-8` を使っている（`apiPost` / `sseSubmitAll` / index.htmlの`executeScan`）。

---

## 2. スプレッドシートのデータ構造

コード.gs 側で `SpreadsheetApp.getActiveSpreadsheet()`（バインド先＝メインの管理用スプレッドシート）と、`CALENDAR_CONFIG.SPREADSHEET_ID`（`1HWZDOIJaQB0S3K4a9hXomJZFqmdznmKywQ2Sc_ON1Fo`）で明示的に開く別スプレッドシートの2系統がある。**実際には同一のスプレッドシートIDを指すよう運用されている前提**（`getActiveSpreadsheet()` を使う関数と `CALENDAR_CONFIG.SPREADSHEET_ID` を使う関数が同じ「公開プロフィール」シートを参照しているため）。

### 2.1 「公開プロフィール」シート（生徒マスタ）
`IDX_PROFILE`（コード.gs 1-17行目）で一元管理：

| 列 | 意味 | 定数 |
|---|---|---|
| A (0) | 生徒ID | `ID` |
| B (1) | 生徒メールアドレス | `STUDENT_EMAIL` |
| C (2) | 保護者メールアドレス | `PARENT_EMAIL` |
| D (3) | 氏名 | `NAME` |
| E (4) | ニックネーム | `NICKNAME` |
| F (5) | 週間目標時間 | `GOAL_HOURS` |
| G (6) | URLトークン（dashboard.htmlの`?token=`と照合） | `TOKEN` |
| H (7) | 目標模試名 | `EXAM_NAME` |
| I (8) | 目標模試日 | `EXAM_DATE` |
| J (9) | （未使用・欠番） | ― |
| K (10) | 生徒個別のGoogleカレンダーID | `CALENDAR_ID` |

**これが唯一の正のマッピング。** 以前は `コード.gs` と `AbsenceCheck.gs` にそれぞれ別の列定義（インデックス基準もバラバラ）が重複しており、それが過去のバグ（トークン列のズレ、カレンダーID列の不一致）の直接原因だった。現在は `AbsenceCheck.gs` 側は独自定義を持たず、`IDX_PROFILE` / `CALENDAR_CONFIG` を直接参照している。

### 2.2 「管理シート」（入退室ログ）
`IDX_LOG`（コード.gs 40-46行目）：A=日付, B=生徒ID, D=入室時刻, E=退室時刻, F=エール数（`CHEERS`）。
- C列は使われていない（IDX_LOGにC相当の定義なし）。
- F列「エール数」は`getStudentStats`が集計して`cheers`としてAPIレスポンスに含めているが、**dashboard.html側にこの値を表示する箇所は現存しない**（`sendCheer()`関数もコード.gs側で「エール機能は廃止されました」という文字列を返すだけの空実装になっている）。旧機能の残骸としてバックエンド側にだけ集計ロジックが残っている状態。

### 2.3 「生徒設定」シート
列：[生徒ID, 模試名, 模試日程]（0,1,2列、`IDX_PROFILE`とは無関係の独自レイアウト）。`getStudentStats`・`saveStudentSettings`が直接ハードコードされた列番号で読み書きする。生徒が dashboard.html の「⚙ 目標模試を自分専用に設定する」から保存すると、ここに1行追記/更新される。シートが存在しない場合は `getStudentStats` 内で自動作成される（`ss.getSheetByName('生徒設定') || ss.insertSheet('生徒設定')`）。

### 2.4 「送信済みログ」シート（欠席確認メール用・AbsenceCheck.gs）
列：[日付, 送信キー, 生徒ID, 生徒名, 予定開始, 送信先]。`ensureSheet`で自動作成。送信キーは `${studentId}_${予定開始日時}` で、同日中の重複送信防止に使う。

### 2.5 「催促送信済みログ」シート（AbsenceCheck.gs）
列：[送信キー, 生徒ID, 生徒名, バージョン, 対象週, 送信先, 送信日時]。Version A/B共通の1シートに両方のログが記録される。送信キーは `A_${studentId}_${週}` / `B_${studentId}_${週}` でバージョンごとに独立。

### 2.6 参考書おすすめ機能用シート
- 「参考書投稿」：[ID, 投稿日時, 生徒ID, 書名, 評価, コメント]
- 「参考書リアクション」：[投稿ID, 生徒ID, 絵文字, 日時]（同一生徒×同一投稿×同一絵文字の組み合わせはトグル式で行ごと追加/削除）

いずれも初回アクセス時に `ensureSheet` で自動作成される（`initBookFeatureSheets()` で手動初期化も可能）。

### 2.7 小テスト進捗用スプレッドシート（別ファイル）
`TEST_SS_ID`（`1uKXnpKeGCuyPpRAryFP7Ou4K4t3SgTNksZFQAhvj45E`）という**別のスプレッドシート**。生徒名（またはID）と同名のシートが生徒ごとに存在し、A列＝参考書名、B〜AD列（1〜30列目）＝チェックボックス（true/false）で各コマの完了状況を表す。生徒のシートが存在しない場合、テンプレートシート（`テンプレート`という名前で存在する前提）を `LockService` で排他制御しながら複製して新規作成する。

---

## 3. コード.gs 関数リファレンス

### エントリポイント
- **`doGet(e)`**：`?token=...&action=...` で呼ばれる。`action` は `stats` / `heatmap` / `bookPosts` / `getCalendarEvents` のいずれか。tokenが無い場合は `index.html`（受付画面）をHTMLとして返す（GASにバインドされたHTML出力。ただし実運用ではGitHub Pages側の`index.html`が使われており、この分岐は事実上使われていない可能性が高い＝旧アーキテクチャの名残）。
- **`doPost(e)`**：JSON body の `action` で分岐。`saveSettings` / `scan` / `postBook` / `reactBook` / `addCalendarEvents` の5種類。

### データ集計
- **`getStudentStats(studentToken)`**：ダッシュボードのメインデータ。`CacheService`で **15分（900秒）キャッシュ**。ログ全行を毎回スキャンする力技実装（生徒数・ログ行数が今後大きく増えると重くなる可能性がある）。
  - 「入室中（未退室）」の生徒については、今週ランキング集計に限り**暫定的に現在時刻までの経過時間を加算**する。ただし未退室が12時間を超えている場合は異常値とみなして加算しない（`liveMs < 12 * 3600000` のガード）。同様の12時間キャップは `getHeatmap` にもある。
  - `calculateStreak`の連続日数判定は「直近の記録日から見て1日以内（＝前日 or 当日）なら連続とみなす」というルール。
  - 週の開始は月曜日固定（`getMonday`）。
- **`getHeatmap(token)`**：直近7日×24時間の在室分数。未退室の場合は「入室時刻+12時間」と「現在時刻」の早い方を終了時刻とみなす。
- **`calculateLevel(totalHours)`**：レベルNからN+1に必要な時間は `5 + N*2` 時間（Lv1→2は7h、Lv2→3は9h...）。
- **`calculateRank(hours)`**：BRONZE(0-10h) → SILVER(10-50h) → GOLD(50-150h) → PLATINUM(150-300h) → SSS MASTER(300h~) の5段階。

### 受付処理
- **`processScan(studentId)`**：QRスキャン（実体はID文字列の読み取り）→入退室トグル。
  - その日の最終データ行を**手動で後方走査**して特定するロジックがある（GASの`getLastRow()`は末尾に空行があると誤動作しやすいため、B列を末尾から遡って実データ行を探している）。
  - 同日中の未退室レコードを探すために、直近100行→300行→900行と**探索範囲を3倍ずつ広げながら最大3回試行**する独自のバックオフ実装（`searchSize *= 3`）。データが少ない日は1回で見つかる想定。
  - 退室時、保護者メール（`notifyParent`）とGoogle Chat通知（`notifyGoogleChat`）を送るが、**両方ともtry/catchで握りつぶしており、失敗してもスキャン自体は成功として返す**（入退室の記録を優先し、通知の失敗でユーザー体験を止めない設計）。

### メール・通知
- **`notifyParent`**：`MailApp.sendEmail`で保護者に入退室を通知（無条件・毎回送信）。
- **`notifyGoogleChat`**：`GCHAT_WEBHOOK_URL`（コード.gs内に直書きされたWebhook URL＋トークン）にPOST。
- **`createDraftsByIds(idsString)`**：スプレッドシートのメニュー（`onOpen`→`showDraftDialog`→`draftDialog.html`）から呼ばれる。カンマ/読点/空白区切りのID一覧を受け取り、各生徒宛にマイページURL案内メールの**下書き**をGmailに作成する（実際の送信はしない＝人間の確認を挟む設計）。

### カレンダー連携
- **`getCalendarEventsForStudent(token, year, month)`**：`doGet`の`getCalendarEvents`アクションから呼ばれる。生徒のトークンから`CALENDAR_ID`列を引き、その生徒個人のGoogleカレンダーから指定月の予定を取得して返す。カレンダーID未設定・取得失敗時は空配列を返す（エラーにしない）。
- **`addCalendarEventsForStudent(token, events)`**：`doPost`の`addCalendarEvents`アクションから呼ばれる。dashboard.htmlの予定入力モーダルから送られた `{date, start, end, memo}[]` を、生徒個人カレンダーに `calendar.createEvent()` で一括登録する。イベントタイトルは `memo` があれば `自習室 - ${memo}`、無ければ `自習室` 固定。

### その他ユーティリティ
- **`ensureSheet(sheetName, headers)`**：シートが無ければヘッダー付きで作成し、`setFrozenRows(1)`する共通ヘルパー。コード.gs / AbsenceCheck.gs 双方から使われる。
- **`findStudentByToken(token)`**：トークン→`{id, name}`変換。`getHeatmap`が使用。
- **`debugWeeklyRanking()` / `debugProfileColumns()`**：手動実行専用のデバッグ用関数（本番フローからは呼ばれない）。`debugProfileColumns()`はシート上の実際のヘッダーと`IDX_PROFILE`の対応を突き合わせてログ出力するので、列がズレた時の調査に便利。

---

## 4. AbsenceCheck.gs 関数リファレンス

同一GASプロジェクト内の別ファイル。**大きく2つの独立した自動送信バッチ処理**が1ファイルにまとまっている。

### 4.1 欠席確認メール（`checkAbsenceAndSendMail`）
- 想定運用：**5分ごとの時間トリガー**（`setupTrigger()`を初回に手動実行して登録）。
- 全生徒をループし、各生徒のGoogleカレンダーから「本日の予定」を取得。予定開始時刻から `CONFIG.ABSENCE_MINUTES`（5分）以上経過していて、かつ当日まだ入室記録がなく、かつ同じキーで未送信なら、生徒＋保護者宛に「来室確認できません」メールを送る。
- 重複送信防止キーは `${studentId}_${予定開始のyyyy/MM/dd HH:mm}`。同じ予定に対しては1日1回しか送らない。

### 4.2 カレンダー未入力の催促メール（`sendReminderVersionA` / `sendReminderVersionB`）
- 想定運用：**毎週月曜8:00の週次トリガー**（`setupReminderTrigger()`で登録）。
- **Version A**：来週（月〜日）の予定が1件も入っていない生徒に送る。
- **Version B**：今週 or 来週のどちらかが空の生徒に送る（Aより広い条件でヒットしやすい）。
- ⚠️ **重要な既知の問題**（コード内に注意コメントあり、AbsenceCheck.gs 409-418行目）：`setupReminderTrigger()` は Version A と B の両方のトリガーを同時に登録してしまう。両方とも「毎週月曜8:00」指定で、GASの時間トリガーは分単位を指定できないためほぼ同時に実行される。**このまま本運用に入れると、条件に一致した生徒には催促メールが2通（A・B）届く。** ユーザー（現在の保守担当者）自身が「どちらを採用するか検討中」と認識しており、本運用前にどちらか一方だけを登録し直す必要がある。**これは開発者本人が把握している既知の未決事項であり、Claude側の推測で処理を変更すべきではない。**

### 4.3 共通ヘルパー
- **`buildRecipients(studentMail, parentMail)`**：生徒＋保護者メールをカンマ区切りにまとめる（保護者メールが空文字/'undefined'なら生徒のみ）。欠席確認・催促メール共通。
- **`sendMailSafely(recipients, subject, body, label, onSuccess)`**：`GmailApp.sendEmail`を try/catch で包み、成功時のみ `onSuccess`コールバック（シートへのログ記録）を呼ぶ共通ヘルパー。3箇所（欠席確認、催促A、催促B）から呼ばれるが、**各呼び出し元固有の「週選択・対象判定」ロジックはあえて統合せず、共通化したのは送信＋ログ記録の定型処理のみ**（統合し過ぎるとVersion A/Bの差分の意味が失われるため、意図的に分離を保っている）。
- **`getSentKeys` / `getReminderSentKeys`**：当日 or 全期間の送信済みキー集合を作るヘルパー。前者は「本日分のみ」、後者は「全期間分」を見る点に違いがある（催促メールは週1回粒度なので当日フィルタが不要なため）。

---

## 5. dashboard.html（生徒用マイページ）仕様

### 5.1 全体構成
1つのHTMLファイルにCSS（`<style>`）とJS（末尾の`<script>`、行3440〜4679）がすべてインラインで書かれた単一ページアプリ。フレームワーク不使用（Vanilla JS）。外部ライブラリはChart.js（CDN）とQRコード生成用の`qrcode.min.js`（同梱ファイル）のみ。

URLの `?token=xxx` パラメータ（`studentToken`、4449行目付近）で生徒を識別する。**このURLは生徒ごとに個別発行され、他人に共有されると誰でもその生徒のデータを閲覧・操作できてしまう**（認証はトークンの知識のみに依存する簡易方式）。

タブは4つ：`HOME` / `TRENDS` / `BOOKS` / `LOG`（`data-tab`属性で切替、`switchTab()`）。TRENDSとBOOKSタブは初回表示時に遅延ロードされる（`heatmapLoaded` / `booksLoaded`フラグ）。

### 5.2 API通信層
- `apiGet(params)`：`GET ${API_ENDPOINT}?...`
- `apiPost(payload)`：`POST ${API_ENDPOINT}`、`Content-Type: text/plain;charset=utf-8`（CORSプリフライト回避のため意図的）
- 上記2つをラップした `fetchStats` / `fetchHeatmap` / `fetchBooks` / `saveStudentSettings` / `apiPostBook` / `apiReactBook` がある。
- **カレンダー関連（`fetchEvents` / `sseSubmitAll`）だけは上記のラッパーを使わず、個別に生の`fetch`を書いている**（実装時期が別だったための不統一。動作に問題はないが、様式は揃っていない）。

### 5.3 ローカルストレージ利用一覧（すべてブラウザ単位・生徒間で共有されない）
| キー | 用途 |
|---|---|
| `sss_ach_unlocked` | 解除済み実績IDの配列（実績の「初回解除」演出を1回だけ出すためのフラグ管理） |
| `sss_missions_YYYY-M-D` | その日の3件のデイリーミッション内容＋手動チェック状態 |
| `sss_mission_all_YYYY-M-D` | その日のミッション全達成演出を出したかどうか |
| `sss_yest_percent` | 週次目標の達成率（ミッション「goal」の前回値比較に使用） |
| `sss_book_today_YYYY-M-D` | その日に参考書投稿済みかどうか（ミッション「book」判定用） |
| `sss_streak_flag` | 連続日数マイルストーン（7/30/100日）の演出を出したかどうか |
| `sss_goal_week_YYYY-M-D` | その週の目標達成演出を出したかどうか（週の月曜日をキーにする） |
| `qr_cache_${studentId}` | QRコードのdata URLキャッシュ（再生成コスト削減） |

### 5.4 実績（ACHIEVEMENTS）とミッション（MISSION_POOL）
- `ACHIEVEMENTS`（3746行目〜）：15種類の固定実績。条件はすべて`getStudentStats`のレスポンスから計算できるクライアント側判定（サーバー側に実績テーブルは無い）。
- `MISSION_POOL`（3800行目〜）：7種類のミッション定義から、日替わりで3件を**日付ベースの疑似乱数シャッフル**（`hashStr`によるシード付きソート、`localStorage`にその日の抽選結果をキャッシュ）で選ぶ。
  - ⚠️ `id: 'test'`（「小テストを1つ進める」）のミッションは `check: d => false` と**常にfalseがハードコードされており、自動達成することはない**（手動チェックのみ可能）。実装が未完のまま残っている可能性が高いが、意図の確認が取れていないため未修正。

### 5.5 予定入力モーダル（Schedule entry / `sse-*` 接頭辞）
`+ 予定を追加` ボタン（`#sse-open-btn`）→ `openModal('sse-modal')`。
- 1件ずつ「日付・開始・終了・メモ」を入力し `+ この予定を追加する`（`sseAddSchedule`）でローカルの`schedules`配列に追加。**この時点ではまだサーバーに送信されない**（複数件まとめて後で一括送信する設計）。
- 繰り返し設定：`なし` / `毎週(4週)` / `曜日指定(2週)` の3モード。「毎週」は同じ曜日で4週分、「曜日指定」は選択した曜日について2週間（14日）分を機械的に展開する。
- `Googleカレンダーに登録する`（`sseSubmitAll`）で `schedules` 配列全体を1回のPOSTで `addCalendarEvents` アクションに送信。成功したら `eventsCache` をクリアして `render()` を呼び、カレンダー表示（週/月ビュー）を即座に再取得・再描画する。

### 5.6 カレンダー表示（週/月ビュー）
- `fetchEvents(year, month)` は月単位でGASから取得し、`eventsCache`（メモリ内、ページ再読み込みで消える）にキャッシュ。
- イベントチップは「時刻（`chip-time`）」と「タイトル（`chip-title`）」を縦に2段表示する構造。これは長いタイトルがグリッドを横に押し広げて崩れる不具合（`min-width:auto`のflexboxデフォルト挙動）を直すために、要約表示（末尾省略）ではなく**全文を折り返して見せる方式**で対応した経緯がある。
- 月表示では1日あたり最大2件までチップを表示し、それ以上は `+N` の集約チップ（`.sss-event-chip.more`）を出す。週表示は件数制限なし（全件縦に並ぶ）。

### 5.7 モーダルの開閉
`openModal(modalId)` / `closeAllModals()` の2関数に統一されている（`settings-modal`・`sse-modal`共通）。以前は予定入力モーダル専用の別実装（独自のoverlay要素・別関数）が並存していたが、重複のため汎用実装に統合済み。

### 5.8 イベント配線
すべて `addEventListener` ベース（`onclick`属性は現在コード全体に1つも残っていない）。静的要素は `setupUIEvents()`（3560行目）に集約。動的に`innerHTML`で再生成される要素（ミッション一覧・参考書リアクションボタン・予定入力の削除ボタン）は、親要素へのイベント委譲（`closest()`＋`data-*`属性読み取り）で処理している。

---

## 6. index.html（受付iPad用スキャン画面）仕様

シンプルな全画面フォーム。タップでオーバーレイを閉じてテキスト入力にフォーカスし、USB/Bluetoothバーコードスキャナ（キーボード入力として認識されるタイプ）でIDを読み取って**Enterキー相当の入力**で送信する想定（`keydown`の`Enter`判定）。

- 入力中に全角文字が混入したら即座に削除する正規表現フィルタがある（スキャナの誤読・IME混入対策と思われる）。
- `setInterval`で1秒ごとに、入力欄からフォーカスが外れていたら強制的に戻す（iPadでの操作中にフォーカスが外れて次のスキャンを取りこぼす事故を防ぐため）。
- APIは `doPost`の`scan`アクションを叩き、`processScan`の結果メッセージをそのまま表示する。

### API_ENDPOINTについて（2026-09-14 修正済み）
`dashboard.html`の`API_ENDPOINT`（3444行目）は最新デプロイURLに更新済み。`index.html`の`API_ENDPOINT`（212行目）も同日中に同じURLへ修正済み。両ファイルとも以下のURLで統一されている：

```
https://script.google.com/macros/s/AKfycbyXbt6myq6hiFODjR4wI9LB_1jUz-JJvsqRsmHlL5VQ4L62KuVgT3q8b9lVarwMGxf0uw/exec
```

**今後GASを再デプロイした際は、`dashboard.html`と`index.html`の両方の`API_ENDPOINT`を新URLに更新すること**（片方だけ更新すると、受付とマイページで異なるバックエンドを向いてしまう）。

---

## 7. draftDialog.html 仕様

GASエディタのスプレッドシート側メニュー（`onOpen()`→「★管理メニュー」→「QRメールの下書き作成」）から `showDraftDialog()` で開かれるモーダルダイアログ。`google.script.run`（GASのHTMLService専用API、`fetch`は使わない）で直接 `createDraftsByIds()` を呼び出す。生徒ID一覧をカンマ区切りで入力すると、各生徒宛のマイページURL案内メールをGmailに**下書きとして**作成する（誤送信防止のため、実送信は人間が下書きを確認してから行う設計）。

---

## 8. 既知の問題・要確認事項（優先度付き）

1. **【既知・運用注意】`setupReminderTrigger()`がVersion A/Bを同時登録する**（4.2章参照）。本運用開始前にどちらか一方に絞る必要あり。コード内にコメント済み。
2. **【軽微・見た目のみ】dashboard.htmlの`<style>`内で、`:root`変数定義とベーススタイル（`*`, `html,body`, `body::before`, `body::after`, `.layout`, `.station-header`〜`.section-header`関連）約300行が、ファイル内で2回（28〜330行目付近と331〜632行目付近）ほぼそのまま重複している。** 値は完全に同一なので後勝ちで上書きされるだけで見た目上の実害は無いが、コードの二重管理状態になっている。過去のセッションで行った「重複CSS統合」作業はプロパティ単位の小さな重複が対象で、このブロック丸ごとの重複は対象外だった（未発見）。整理するなら一方を丸ごと削除するだけで済む単純な修正。
3. **【設計メモ】`エール（cheers）`機能はフロントエンドから撤去済みだが、バックエンド（`getStudentStats`の`cheers`集計、`管理シート`F列読み取り）は残っている。** 表示先が無いだけでAPIレスポンスには今も含まれる。実害はないが、将来的に「使われていないフィールド」として整理対象になりうる。
4. **【運用メモ】`TARGET_COMMON`（共通テスト日）・`TARGET_NATIONAL`（国立二次試験日）はdashboard.html内に西暦日付でハードコーディングされている**（現在値：2027/01/16, 2027/02/25）。年度が変わるたびに手動更新が必要。
5. **【設計メモ】`MISSION_POOL`の`id:'test'`ミッションは条件が常に`false`で、自動達成することがない**（手動チェックのみ）。未完成実装の可能性があるが、意図未確認のため放置中。
6. **【設計メモ】`window._sssInitDone`フラグ**（初期データ読み込み後1.5秒で`true`になる）は、ページ読み込み直後に実績・ストリークの「解除演出」が誤って再生されるのを防ぐためのタイミング制御と推測されるが、実装者の意図は未確認。
7. **【秘匿情報】`GCHAT_WEBHOOK_URL`（Google Chat Webhook・トークン付きURL）、`CALENDAR_CONFIG.SPREADSHEET_ID`、`TEST_SS_ID`などがコード.gs内に平文でハードコーディングされている。** GASの性質上一般的ではあるが、リポジトリがpublicになった場合は要注意（現状GitHubリポジトリの公開設定は本ドキュメントの範囲外につき未確認）。

---

## 9. 操作手順（半年後の自分向け・そのまま実行できるチェックリスト）

このセクションだけは「取説」として使えるよう、コードを読み返さなくても実行できる手順に絞って書く。

### 9.1 GAS Web Appを再デプロイする手順
1. GASエディタで コード.gs / AbsenceCheck.gs を編集・保存する。
2. 右上「デプロイ」→「デプロイを管理」→ 既存デプロイの鉛筆アイコン →「バージョン」で「新しいバージョン」を選択 →「デプロイ」。
   - 既存デプロイを「新しいバージョン」で更新する限り、Web App のURLは変わらない。URLが変わるのは「新しいデプロイ」を新規作成した場合のみ（基本的には新規作成しなくてよい）。
3. もしURLが変わった場合（新規デプロイを作った場合）は、`dashboard.html`（3444行目）と`index.html`（212行目）の**両方**の`API_ENDPOINT`を新URLに更新する（片方だけ更新すると受付とマイページが別バックエンドを向く。9.4章のトラブル表も参照）。
4. `git add` → `commit` → `push`（SSHエイリアス `github.com.arbeiten` 経由）でGitHub Pagesに反映する。
5. GitHub Pagesへの反映には数分ラグが出ることがある。反映後、実際にdashboard.html/index.htmlをブラウザで開いて動作確認する。

### 9.2 新入生を追加する手順
1. 「公開プロフィール」シートに1行追加する。
2. 各列を埋める（列の意味は2.1章の表を参照）：ID／生徒メール／保護者メール／氏名／ニックネーム／週間目標時間／URLトークン（他の生徒と重複しないランダム文字列を自分で発行して入力）／目標模試名・日（空欄でも可）／J列は空欄のまま／カレンダーID（その生徒用に用意したGoogleカレンダーのIDを控えて入力）。
3. GASエディタのスプレッドシート側メニュー「★管理メニュー」→「QRメールの下書き作成」を開き、新入生のIDを入力して実行する（`createDraftsByIds`）。→ Gmailの下書きフォルダにマイページURL案内メールが作成される。
4. 下書きの内容を確認してから手動で送信する（自動送信はしない設計）。

### 9.3 トリガーの初回登録（新しいGASプロジェクトに移行した場合など）
1. GASエディタで関数`setupTrigger`を選択して実行（▶ボタン）。→ 欠席確認メールが5分ごとに走るようになる。
2. `setupReminderTrigger`を実行する前に、催促メールをVersion A・Bどちらで運用するか決めること（8章-1参照、現時点では未決定）。決めずに実行すると両方登録されて二重送信になる。実行してしまった場合は、GASエディタ左メニュー「トリガー」から片方を手動削除する。

### 9.4 よくある不具合の切り分け表
| 症状 | まず確認する箇所 |
|---|---|
| 生徒のマイページで「URLが正しくありません」と出る | URLの`?token=`の値と「公開プロフィール」G列が一致しているか |
| 生徒のマイページで「無効なURLです。管理者に連絡してください」と出る | 同上（サーバー側`getStudentStats`の`targetId`未検出エラー） |
| 受付iPadでスキャンしてもエラー・無反応 | `dashboard.html`と`index.html`の`API_ENDPOINT`が同じURLか（9.1章）、GAS側のデプロイが有効か |
| カレンダーの予定が生徒のページに反映されない | 「公開プロフィール」K列（カレンダーID）が正しいか、そのGoogleカレンダーへGASのアカウントがアクセスできるか |
| 催促・欠席確認メールが二重に届く | `setupReminderTrigger()`でVersion A/B両方のトリガーが有効になっていないか（GASエディタ「トリガー」画面で確認、8章-1参照） |
| データの列がズレて表示される・変な値が出る | GASエディタで`debugProfileColumns()`を手動実行し、ログでシートの実際の列と`IDX_PROFILE`の対応を突き合わせる |

### 9.5 年次メンテナンス
- `TARGET_COMMON` / `TARGET_NATIONAL`（dashboard.html 3445-3446行目）：共通テスト・国立二次試験の目標日。年度が変わったら西暦日付を手動で書き換える。

### 9.6 その他の運用情報
- **GitHub Pages**：`dashboard.html` / `index.html` を含むこのリポジトリを静的ホスティング。SSH経由でのpush用に、`~/.ssh/config`に `github.com.arbeiten` というホストエイリアスと専用鍵を設定済み（他アカウント用のデフォルト`github.com`設定と分離）。手順は9.1章参照。

---

## 10. このドキュメントの作成方針について

このリポジトリは複数の開発者が関わってきており、現在の保守担当者（このドキュメントの依頼者）は「自分が書いていないコードの意図を推測でコメント化することはしたくない」という方針を明確にしている。そのためコード本体へのコメント追加は必要最小限（`AbsenceCheck.gs`の`setupReminderTrigger()`のみ、本人が意図を把握している箇所）に留めており、本ドキュメントはコードを変更せずに**外部から見た仕様・挙動・既知の問題を客観的に記述する**形でその方針を補っている。今後このリポジトリを引き継ぐ人（人間・AIエージェント問わず）は、まずこのドキュメントで全体像を掴んだ上で、個々のロジックの「なぜそうなっているか」を変更する前には保守担当者に確認することを推奨する。
