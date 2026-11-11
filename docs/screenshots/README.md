# 説明用の図について

`form-transfer-demo.gif` と `form-transfer-before-after.png` は、デモ用アカウントでの **実操作画面** です。顧客の本番環境ではありません。通知先のメールアドレスやプロフィール画像などの個人情報は伏せています。

それ以外の PNG は、**実際の Google スプレッドシートや Gmail の画面を撮影したものではありません。** 列構成と処理前後の違い、トリガーに選ぶ項目、メール本文の例を、説明用に組んだ図です。

ダミーの名前と `example.com` のアドレスだけを使っています。実在の顧客情報はありません。

図の元になる HTML は `docs/screenshots/mocks/` にあります。再生成する場合のコマンド例（リポジトリのルートで）:

```bash
python3 docs/screenshots/render.py
```
