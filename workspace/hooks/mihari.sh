#!/bin/zsh
# 芯 core：記憶の見張り（会話の始まりに・道具なし）。問題がある時だけ一言を出す。
# 1件 本文15行（約1.5KB）まで・1件1事実（「なぜ：」は1つ）・目次40行まで・手順書と hook がそろっているか。
export LC_ALL=ja_JP.UTF-8
cd "${CLAUDE_PROJECT_DIR:-$PWD}" 2>/dev/null || exit 0
out=()
for f in 記憶/*.md(N); do
  [[ $f == */目次.md ]] && continue
  body=$(awk 'BEGIN{h=0} /^---$/{h++; next} h>=2||h==0' "$f" | grep -v '^[[:space:]]*$')
  n=$(print -r -- "$body" | grep -c .); sz=$(wc -c < "$f" | tr -d ' ')
  w=$(grep -c '^なぜ[：:]' "$f")
  (( n > 15 || sz > 1536 )) && out+="「${f:t:r}」が長い（本文 ${n}行・${sz}バイト）→ 割る"
  (( w > 1 )) && out+="「${f:t:r}」に事実が ${w}つ → 割る"
  grep -q "^- \[.*\](${f:t})" 記憶/目次.md 2>/dev/null || out+="「${f:t:r}」が目次に無い → 1行足す"
done
[ -f 記憶/目次.md ] && (( $(grep -c '^- \[' 記憶/目次.md) > 40 )) && out+="目次が40行を超えた → 束ねる"
for s in 思い出す 記憶に残す 記憶の手入れ; do [ -f .claude/skills/$s/SKILL.md ] || out+="手順書「$s」が無い"; done
(( ${#out} == 0 )) && exit 0
print "⚠️ 記憶の見張り：${(j:／:)out}。手順書「記憶の手入れ」で直す案を出してから、今日の話に入る。"
exit 0
