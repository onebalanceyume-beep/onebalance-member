# キッズ体験予約 GAS セットアップ
1. ジムアカウント（one.balance.yume@gmail.com）で https://script.google.com → 新しいプロジェクト → 名前「キッズ体験予約」
2. コード.gs を全消しして gas/kids-taiken.gs を貼る → 保存
3. 関数「setupOnce」を選んで実行 → 権限を許可（カレンダー/スプレッドシート/メール）
4. デプロイ → 新しいデプロイ → ウェブアプリ → 実行ユーザー=自分 / アクセス=全員 → デプロイ
5. 出た /exec のURLを kids-taiken.html の `var KIDS_GAS = "";` に貼る
