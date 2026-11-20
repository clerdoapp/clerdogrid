#!/usr/bin/env bash
# Grid storing workbooks in Locker, end to end: both servers and one mock Clerdo ID.
#   bash server/dev/locker-test.sh  [LOCKER_ROOT=/path/to/locker]
set -uo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
LOCKER_ROOT="${LOCKER_ROOT:-$ROOT/../../locker}"
RUN="$ROOT/dev/.run-locker"; mkdir -p "$RUN"
export GRID_RUN_DIR="$RUN" GRID_SQLITE="$RUN/grid.sqlite" GRID_IDP_STATE="$RUN/idp.json"
export GRID_APP_PORT=8180 GRID_IDP_PORT=8190 GRID_LOCKER_URL=http://127.0.0.1:8280
export LOCKER_RUN_DIR="$RUN/locker" LOCKER_SQLITE="$RUN/locker.sqlite" LOCKER_STORAGE="$RUN/locker-data"
export LOCKER_APP_PORT=8280 LOCKER_IDP_PORT=8290 LOCKER_IDP=http://127.0.0.1:8190 LOCKER_IDP_STATE="$RUN/locker-idp.json"
rm -rf "$RUN"/*.sqlite* "$RUN"/*.json "$LOCKER_STORAGE"; mkdir -p "$LOCKER_RUN_DIR"
bash "$ROOT/dev/start.sh" >/dev/null || { echo "grid start failed"; exit 1; }
bash "$LOCKER_ROOT/server/dev/start.sh" >/dev/null || { echo "locker start failed"; bash "$ROOT/dev/start.sh" stop >/dev/null; exit 1; }
stop() { bash "$ROOT/dev/start.sh" stop >/dev/null; bash "$LOCKER_ROOT/server/dev/start.sh" stop >/dev/null; }
trap stop EXIT
APP="http://127.0.0.1:$GRID_APP_PORT"; IDP="http://127.0.0.1:$GRID_IDP_PORT"; LK="http://127.0.0.1:$LOCKER_APP_PORT"
J1="$RUN/jar1.txt"; rm -f "$J1"
PASS=0; FAIL=0
ok()   { PASS=$((PASS+1)); echo "  ok   $1"; }
bad()  { FAIL=$((FAIL+1)); echo "  FAIL $1 ${2:-}"; }
check(){ if [ "$2" = "$3" ]; then ok "$1"; else bad "$1" "(got '$2', want '$3')"; fi; }
jq_()  { php -r '$d=json_decode(stream_get_contents(STDIN),true); $p=explode(".",$argv[1]); foreach($p as $k){ if($k==="") continue; if(is_array($d)&&array_key_exists($k,$d)) $d=$d[$k]; else { $d=null; break; } } if(is_bool($d)) echo $d?"true":"false"; elseif(is_null($d)) echo "null"; elseif(is_scalar($d)) echo $d; else echo json_encode($d);' "$1"; }
# req JAR METHOD PATH [DATA|@FILE] [extra curl args...] → sets BODY, CODE
req() {
  local jar=$1 m=$2 p=$3; shift 3
  local data=''
  if [ $# -gt 0 ] && [[ "$1" != -* ]]; then data=$1; shift; fi
  local out
  if [ -n "$data" ]; then
    out=$(curl -s -b "$jar" -c "$jar" -X "$m" "$APP$p" -H 'Content-Type: application/json' --data-binary "$data" "$@" -w $'\n%{http_code}')
  else
    out=$(curl -s -b "$jar" -c "$jar" -X "$m" "$APP$p" "$@" -w $'\n%{http_code}')
  fi
  CODE=${out##*$'\n'}; BODY=${out%$'\n'*}
}
mreq() { local jar=$1 csrf=$2; shift 2; req "$jar" "$@" -H "X-CSRF-Token: $csrf" -H "Origin: $APP"; }
login() { # jar user
  local link
  link=$(curl -s -b "$1" -c "$1" -L "$APP/auth/login?next=/?p=after" | grep -o "/authorize?[^\"]*user=$2" | head -1 | sed 's/&amp;/\&/g')
  curl -s -L -b "$1" -c "$1" "$IDP$link" -o /dev/null -w '%{http_code} %{url_effective}'
}
# A minimal .xlsx with one value in A1, as base64.
xlsx() {
  php -r '
    $t = tempnam(sys_get_temp_dir(), "x"); $z = new ZipArchive(); $z->open($t, ZipArchive::OVERWRITE);
    $z->addFromString("[Content_Types].xml", "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
    $z->addFromString("xl/workbook.xml", "<workbook/>");
    $z->addFromString("xl/worksheets/sheet1.xml", "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>" . $argv[1] . "</v></c></row></sheetData></worksheet>");
    $z->close(); echo base64_encode(file_get_contents($t)); unlink($t);' "$1"
}
# Writes a JSON body file: {"name":..., "workbook":..., [extra]}
body() { php -r 'echo json_encode(array_merge(["workbook" => $argv[1]], json_decode($argv[2], true)));' "$1" "$2" > "$RUN/body.json"; }


# lk METHOD PATH [curl args...] — Locker as the same user, with a token from the same Clerdo ID
AT=$(curl -s -X POST "$IDP/_dev/token" -H 'Content-Type: application/json' -d '{"user":"u-hamzeh"}' | jq_ access_token)
lk() { local m=$1 p=$2; shift 2; local out; out=$(curl -s -X "$m" "$LK$p" -H "Authorization: Bearer $AT" "$@" -w $'\n%{http_code}'); CODE=${out##*$'\n'}; BODY=${out%$'\n'*}; }
a1() { php -r '$t=tempnam(sys_get_temp_dir(),"x"); file_put_contents($t, base64_decode(stream_get_contents(STDIN))); $z=new ZipArchive(); $z->open($t); preg_match("#<v>(.*?)</v>#", $z->getFromName("xl/worksheets/sheet1.xml"), $m); echo $m[1] ?? "?"; unlink($t);'; }
lka1() { curl -s "$LK/api/items/$1/content" -H "Authorization: Bearer $AT" -o "$RUN/dl.xlsx"; fa1 < "$RUN/dl.xlsx"; }
fa1() { php -r '$t=tempnam(sys_get_temp_dir(),"x"); file_put_contents($t, stream_get_contents(STDIN)); $z=new ZipArchive(); $z->open($t); preg_match("#<v>(.*?)</v>#", $z->getFromName("xl/worksheets/sheet1.xml"), $m); echo $m[1] ?? "?"; unlink($t);'; }

WB1=$(xlsx 1); WB2=$(xlsx 2); WB3=$(xlsx 3); WB4=$(xlsx 4)
login "$J1" u-hamzeh >/dev/null
req "$J1" GET /api/me; CSRF=$(echo "$BODY" | jq_ csrf)

echo "== a new workbook goes to Locker"
body "$WB1" '{"name":"Budget / 2026"}'
mreq "$J1" "$CSRF" POST /api/projects "@$RUN/body.json"
check "create 201" "$CODE" 201
check "stored in Locker" "$(echo "$BODY" | jq_ project.storage)" locker
PID=$(echo "$BODY" | jq_ project.id); LID=$(echo "$BODY" | jq_ project.lockerId)
lk GET "/api/items/$LID"
check "file name in Locker" "$(echo "$BODY" | jq_ item.name)" "Budget - 2026.xlsx"
check "saved by Grid" "$(echo "$BODY" | jq_ item.app)" grid
FOLDER=$(echo "$BODY" | jq_ item.parent)
lk GET "/api/items/$FOLDER"
check "inside the Grid folder" "$(echo "$BODY" | jq_ item.name)" Grid
check "Locker has the content" "$(lka1 $LID)" 1

echo "== saves become Locker versions"
body "$WB2" '{"baseRevision":1}'
mreq "$J1" "$CSRF" PUT "/api/projects/$PID" "@$RUN/body.json"
check "save 200" "$CODE/$(echo "$BODY" | jq_ project.revision)" "200/2"
check "Locker has the new content" "$(lka1 $LID)" 2
lk GET "/api/items/$LID/versions"
check "Locker kept the first version" "$(echo "$BODY" | jq_ versions.0.rev)" 1
req "$J1" GET "/api/projects/$PID/revisions"
check "Grid lists the earlier revision" "$(echo "$BODY" | jq_ revisions.0.revision)" 1
mreq "$J1" "$CSRF" POST "/api/projects/$PID/revisions/1/restore" ''
check "restore from Locker's version" "$CODE/$(echo "$BODY" | jq_ workbook | a1)" "200/1"
check "restore reached Locker" "$(lka1 $LID)" 1

echo "== a change made in Locker is picked up"
req "$J1" GET "/api/projects/$PID"; GREV=$(echo "$BODY" | jq_ project.revision)
lk GET "/api/items/$LID"; LREV=$(echo "$BODY" | jq_ item.rev)
echo -n "$WB4" | base64 -d > "$RUN/wb4.xlsx"
lk PUT "/api/items/$LID/content" -H "If-Match: \"$LREV\"" -H 'Content-Type: application/octet-stream' --data-binary @"$RUN/wb4.xlsx"
check "edited in Locker" "$CODE" 200
body "$WB3" "{\"baseRevision\":$GREV}"
mreq "$J1" "$CSRF" PUT "/api/projects/$PID" "@$RUN/body.json"
check "Grid save over it → conflict" "$CODE/$(echo "$BODY" | jq_ workbook | a1)" "409/4"
NREV=$(echo "$BODY" | jq_ project.revision)
mreq "$J1" "$CSRF" PUT "/api/projects/$PID" "@$RUN/body.json" >/dev/null
body "$WB3" "{\"baseRevision\":$NREV}"
mreq "$J1" "$CSRF" PUT "/api/projects/$PID" "@$RUN/body.json"
check "keep mine after the conflict" "$CODE" 200
check "Locker has the kept version" "$(lka1 $LID)" 3
LREV2=$(curl -s "$LK/api/items/$LID" -H "Authorization: Bearer $AT" | jq_ item.rev)
lk PUT "/api/items/$LID/content" -H "If-Match: \"$LREV2\"" -H 'Content-Type: application/octet-stream' --data-binary @"$RUN/wb4.xlsx"
req "$J1" GET "/api/projects/$PID"
check "opening picks up the Locker change" "$(echo "$BODY" | jq_ workbook | a1)" 4

echo "== rename, share, delete"
mreq "$J1" "$CSRF" PATCH "/api/projects/$PID" '{"name":"Q4 plan"}'
lk GET "/api/items/$LID"
check "rename reaches Locker" "$(echo "$BODY" | jq_ item.name)" "Q4 plan.xlsx"
mreq "$J1" "$CSRF" POST "/api/projects/$PID/share" '{"mode":"view"}'; TOK=$(echo "$BODY" | jq_ token)
req "$RUN/none.txt" GET "/api/share/$TOK"
check "share link still opens" "$(echo "$BODY" | jq_ workbook | a1)" 4
mreq "$J1" "$CSRF" DELETE "/api/projects/$PID" ''
check "delete 204" "$CODE" 204
lk GET "/api/items/$LID"
check "file goes to the Locker Bin" "$([ "$(echo "$BODY" | jq_ item.trashedAt)" != null ] && echo yes)" yes

echo "== Locker unreachable: the save is kept and sent later"
body "$WB1" '{"name":"Offline"}'
mreq "$J1" "$CSRF" POST /api/projects "@$RUN/body.json"; P2=$(echo "$BODY" | jq_ project.id); L2=$(echo "$BODY" | jq_ project.lockerId)
bash "$LOCKER_ROOT/server/dev/start.sh" stop >/dev/null
body "$WB2" '{"baseRevision":1}'
mreq "$J1" "$CSRF" PUT "/api/projects/$P2" "@$RUN/body.json"
check "save still works" "$CODE/$(echo "$BODY" | jq_ project.storage)" "200/pending"
bash "$LOCKER_ROOT/server/dev/start.sh" >/dev/null
req "$J1" GET "/api/projects/$P2"
check "back in Locker on open" "$(echo "$BODY" | jq_ project.storage)" locker
check "Locker caught up" "$(lka1 $L2)" 2

echo "== workbooks from before Locker move in"
php -r '$db=new PDO("sqlite:".$argv[1]); $db->exec("UPDATE grid_projects SET locker_id=NULL, locker_rev=NULL, locker_pending=0 WHERE id=\x27".$argv[2]."\x27");' "$GRID_SQLITE" "$P2"
req "$J1" GET /api/projects
check "listed as not in Locker yet" "$(echo "$BODY" | jq_ projects.0.storage)" pending
req "$J1" GET "/api/projects/$P2"
check "moved into Locker on open" "$(echo "$BODY" | jq_ project.storage)" locker
[ "$(echo "$BODY" | jq_ project.lockerId)" != "$L2" ] && ok "a new Locker file" || bad "a new Locker file"

echo
echo "passed: $PASS   failed: $FAIL"
[ "$FAIL" -eq 0 ]
