# orbit-core

Claude Code の作業場に「覚える・思い出す・見張る」仕組みを足す、手順書と hook のパックです。道具は何も入れません（Mac に元からある zsh・grep・awk・sed だけ）。

## 中身
| 置き場 | 何をするか |
|---|---|
| workspace/hooks/omoidasu.sh | 話しかけるたびに、記憶/ と 資料/ から関係しそうな物の題名を3件まで出す |
| workspace/hooks/mihari.sh | 会話の始まりに、長すぎる記憶・1件に2つ以上の事実・目次に無い記憶・手順書の欠けを一言で知らせる |
| workspace/settings.json | 上の2つを Claude Code の hook に付ける行（芯がまだ無い時は「芯を更新して」と案内する） |
| workspace/skills/芯を更新する/ | 新しい版を取り込む手順書 |
| gas/ | 台帳のスクリプト。当て方は gas/当てる.sh と gas/CHANGELOG.md |

## 使い方
1. 作業場の `.claude/settings.json` に workspace/settings.json の hook の行を入れ、`.claude/skills/` に workspace/skills/ を入れる
2. Claude Code に「芯を更新して」と言う。最新の札を取り込み、hook の中身を `.orbit/core/` に置き、当てた札を `.orbit/版.json` に残す
3. 次の会話から、思い出す・見張るが効く

記憶の形：`記憶/<名前>.md` に1件1事実（頭に name と 説明）・本文15行（約1.5KB）まで・目次 `記憶/目次.md` は40行まで。

## 版
札は日付（core-2026.10.30.2・同じ日の2回目は .2）。変えた所は CHANGELOG.md に1版1行。許可の画面が出る版はそこに書きます。
