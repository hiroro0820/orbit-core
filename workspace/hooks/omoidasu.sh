#!/bin/zsh
# 芯 core：想起の hook（発言のたび）。Mac に元からある zsh だけ・道具を入れない。
# 発言と 記憶/ の name・説明（資料/ はファイル名）で、漢字・カタカナの2字の重なりを数え、上位3件の題名だけを出す。
export LC_ALL=ja_JP.UTF-8
cd "${CLAUDE_PROJECT_DIR:-$PWD}" 2>/dev/null || exit 0
p=$(sed -E -n 's/.*"prompt"[[:space:]]*:[[:space:]]*"(([^"\\]|\\.)*)".*/\1/p' | head -c 2000)
[ -z "$p" ] && exit 0
typeset -A score title
setopt extendedglob
pairs() { local s=$1 i a b; for (( i=1; i<${#s}; i++ )); do a=${s[i]}; b=${s[i+1]}; [[ $a == [一-龥ァ-ヶー々] && $b == [一-龥ァ-ヶー々] ]] && print -r -- "$a$b"; done | sort -u }
for f in 記憶/*.md(N) 資料/*.md(N); do
  [[ $f == */目次.md || $f == */_ひな形.md || $f == */README.md ]] && continue
  n=${f:t:r}; d=""
  [[ $f == 記憶/* ]] && d=$(sed -n -E '1,8s/^(name|説明):[[:space:]]*//p' "$f" | tr '\n' ' ')
  s=0; for g in $(pairs "$n $d"); do [[ $p == *$g* ]] && (( s++ )); done
  (( s >= 2 )) && { score[$f]=$s; title[$f]="$n｜${d[1,60]}"; }
done
(( ${#score} == 0 )) && exit 0
print "関係しそうな記憶・資料（題名だけ。要る物だけ開く）："
for f in ${(k)score}; do print -r -- "${score[$f]} $f"; done | sort -rn | head -3 | while read s f; do print -r -- "- ${title[$f]}｜$f"; done
exit 0
