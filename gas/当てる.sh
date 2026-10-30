#!/bin/zsh
# お客様の作業場が、Orbit の新しい版を自分の台帳のスクリプトに当てる。使い方：作業場の一番上で zsh <取り込んだ gas>/当てる.sh <スクリプトの ID>
# 会社用の部分は作業場の custom/orbit-gas.js（無ければ芯の既定のまま）。更新では上書きしない
# ID は台帳の「Orbit の情報」タブにある。clasp はお客様の Google でログイン済み（最初の1回だけ「許可」を押す）
# 当てた後は、次の振り分け（15分以内）で「Orbit の情報」の版がこの版に変わる
set -e
[ -n "$1" ] || { echo "スクリプトの ID を渡してください（台帳の「Orbit の情報」タブ）"; exit 1; }
H="${0:A:h}"
(cd "$H" && node test.js >/dev/null) || { echo "テストが通らないので当てません"; exit 1; }
D=$(mktemp -d)
cp "$H/Code.js" "$H/台帳つき/メニュー.js" "$H/台帳つき/つなぐ.html" "$H/台帳つき/appsscript.json" "$D/"
C="${ORBIT_CUSTOM:-$PWD/custom/orbit-gas.js}"
[ -f "$C" ] && { node -e "new Function(require('fs').readFileSync(process.argv[1],'utf8'))" "$C" || { echo "custom/orbit-gas.js の書き方が崩れているので当てません"; exit 1; }; cp "$C" "$D/custom.js"; echo "会社用の部分も一緒に置く：$C"; }
printf '{"scriptId":"%s","rootDir":""}\n' "$1" > "$D/.clasp.json"
cd "$D" && npx -y @google/clasp push --force && echo "当てた：$(grep -oE "ORBIT_版 = '[0-9.-]+'" Code.js)"
