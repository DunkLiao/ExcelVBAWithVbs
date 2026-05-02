---
name: git-commit-push
description: 執行 git commit 並 push 到遠端。列出變更、詢問 commit 訊息（若未提供），然後 add、commit 並 push。
---

# Git Commit & Push

執行以下步驟：

1. 執行 `git status` 確認目前有哪些變更
2. 若沒有任何變更，告知使用者並結束
3. 若使用者沒有提供 commit 訊息，根據變更內容自動產生一個簡短的繁體中文訊息
4. 執行 `git add .` 暫存所有變更
5. 執行 `git commit -m "<訊息>"` 提交，commit 訊息結尾必須加上以下 trailer：
   ```
   Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>
   ```
6. 執行 `git push` 推送到遠端
7. 確認成功後回報結果
