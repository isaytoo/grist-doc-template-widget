#!/bin/sh
# Tests du widget dans un vrai Chrome (l'éditeur Jodit ne tourne pas sous jsdom).
# La page de test est le vrai index.html et le vrai widget.js, avec un faux Grist ;
# un « pilote » y joue un scénario et écrit ses constats dans la page.
#
#   sh tests/tout.sh
#
# pilote-nouveautes.js   bordures invisibles, mise en forme des variables
# pilote-resolution.js   fusion de modèles sans mise en forme : la sortie doit rester
#                        celle de reference-resolution.txt, produite avant la v1
T=$(cd "$(dirname "$0")" && pwd); R=$(dirname "$T")
CHROME=${CHROME:-/Applications/Google Chrome.app/Contents/MacOS/Google Chrome}
[ -x "$CHROME" ] || { echo "Chrome introuvable : définir CHROME=/chemin/vers/chrome"; exit 2; }
TMP=$(mktemp -d); trap 'rm -rf "$TMP"' EXIT
jouer() {
  python3 "$T/fabrique.py" "$R" "$R/widget.js" "$TMP/page.html" "$T/$1" || return 1
  "$CHROME" --headless=new --disable-gpu --window-size=1400,900 --virtual-time-budget=14000 --dump-dom "file://$TMP/page.html" 2>/dev/null \
    | grep -o 'RES{.*}FIN' | head -1 | sed 's/^RES//;s/FIN$//' | python3 -c "
import sys,json,html
t=sys.stdin.read()
if not t.strip(): print('KO  le pilote n a rien rendu'); sys.exit(0)
for k,v in json.loads(html.unescape(t)).items(): print('%-36s %s' % (k, v))"
}
ko=0
jouer pilote-nouveautes.js > "$TMP/n.txt"; cat "$TMP/n.txt" | cut -c1-150
grep -q '^KO' "$TMP/n.txt" && ko=1; [ -s "$TMP/n.txt" ] || ko=1
jouer pilote-resolution.js > "$TMP/r.txt"
if diff -q "$TMP/r.txt" "$T/reference-resolution.txt" >/dev/null; then echo "OK  fusion sans mise en forme : $(wc -l < "$TMP/r.txt" | tr -d ' ') sorties identiques à la référence"
else echo "KO  la fusion diffère de la référence :"; diff "$TMP/r.txt" "$T/reference-resolution.txt" | cut -c1-200 | head -12; ko=1; fi
[ $ko -eq 0 ] && echo "Tout passe." || echo "ÉCHEC."
exit $ko
