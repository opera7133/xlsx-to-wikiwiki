# xlsx-to-wikiwiki

`ryceam.xlsx`を受け取り、wikiwiki.jp REST APIで更新するスクリプト。

## 使い方

依存関係のインストール：

```bash
bun install
```

`.env`の作成：

```env
WIKI_NAME=<wiki-name>
PASSWORD=<wiki-password>
```

実行：

```bash
bun run index.ts
```
