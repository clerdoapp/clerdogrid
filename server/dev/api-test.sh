#!/usr/bin/env bash
# End-to-end test of the cloud API against the dev stack (sqlite, mock IdP). Starts and stops its own servers.
#   bash server/dev/api-test.sh
set -uo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
RUN="$ROOT/dev/.run"
mkdir -p "$RUN"
export GRID_RUN_DIR="$RUN"
export GRID_SQLITE="$RUN/grid-test.sqlite"
export GRID_IDP_STATE="$RUN/grid-idp-test.json"
export GRID_SHARE_RPM=30
rm -f "$GRID_SQLITE" "$GRID_SQLITE-wal" "$GRID_SQLITE-shm" "$GRID_IDP_STATE"
GRID_EVERYONE_PRO=0 bash "$ROOT/dev/start.sh" >/dev/null || { echo "start failed"; exit 1; }
trap 'bash "$ROOT/dev/start.sh" stop >/dev/null' EXIT

APP="http://127.0.0.1:${GRID_APP_PORT:-8080}"
IDP="http://127.0.0.1:${GRID_IDP_PORT:-8090}"
J1="$RUN/jar1.txt"; J2="$RUN/jar2.txt"; rm -f "$J1" "$J2"
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

WB1=$(xlsx 1); WB2=$(xlsx 2); WB3=$(xlsx 3)

echo "== auth"
R=$(login "$J1" u-hamzeh)
check "sign-in returns to the page it started on" "$R" "200 $APP/?p=after"
req "$J1" GET /api/me
check "/api/me 200" "$CODE" 200
check "me.user.id" "$(echo "$BODY" | jq_ user.id)" u-hamzeh
CSRF=$(echo "$BODY" | jq_ csrf)
[ ${#CSRF} -eq 64 ] && ok "csrf token present" || bad "csrf token"
req "$RUN/none.txt" GET /api/me
check "signed out → 401" "$CODE" 401

echo "== csrf"
body "$WB1" '{"name":"x"}'
req "$J1" POST /api/projects @"$RUN/body.json" -H "Origin: $APP"
check "missing CSRF → 403" "$CODE" 403
req "$J1" POST /api/projects @"$RUN/body.json" -H "X-CSRF-Token: $CSRF" -H "Origin: https://evil.example"
check "foreign Origin → 403" "$CODE" 403

echo "== create"
body "$WB1" '{"name":"Budget 2026","sheets":2}'
mreq "$J1" "$CSRF" POST /api/projects @"$RUN/body.json"
check "create 201" "$CODE" 201
PID=$(echo "$BODY" | jq_ project.id)
check "create: revision 1" "$(echo "$BODY" | jq_ project.revision)" 1
check "create: sheets 2" "$(echo "$BODY" | jq_ project.sheets)" 2
check "create: name" "$(echo "$BODY" | jq_ project.name)" "Budget 2026"
mreq "$J1" "$CSRF" POST /api/projects '{"name":"x"}'
check "create without workbook → 422" "$CODE" 422
mreq "$J1" "$CSRF" POST /api/projects '{"name":"x","workbook":"not base64!!"}'
check "invalid base64 → 422" "$CODE" 422
NOTZIP=$(printf 'hello this is not a zip file at all......' | base64 -w0)
mreq "$J1" "$CSRF" POST /api/projects "{\"name\":\"x\",\"workbook\":\"$NOTZIP\"}"
check "not a zip → 422" "$CODE" 422
EMPTYZIP=$(php -r '$t=tempnam(sys_get_temp_dir(),"z");$z=new ZipArchive();$z->open($t,ZipArchive::OVERWRITE);$z->addFromString("a.txt","a");$z->close();echo base64_encode(file_get_contents($t));')
mreq "$J1" "$CSRF" POST /api/projects "{\"name\":\"x\",\"workbook\":\"$EMPTYZIP\"}"
check "zip without workbook parts → 422" "$CODE" 422

echo "== read"
req "$J1" GET "/api/projects/$PID"
check "get 200" "$CODE" 200
check "get: workbook round-trips byte for byte" "$(echo "$BODY" | jq_ workbook)" "$WB1"
req "$J1" GET /api/projects
check "list has it" "$(echo "$BODY" | jq_ projects.0.id)" "$PID"

echo "== save and conflicts"
body "$WB2" '{"baseRevision":1,"sheets":1,"name":"Budget"}'
mreq "$J1" "$CSRF" PUT "/api/projects/$PID" @"$RUN/body.json"
check "save 200" "$CODE" 200
check "save: revision 2" "$(echo "$BODY" | jq_ project.revision)" 2
check "save: renamed" "$(echo "$BODY" | jq_ project.name)" "Budget"
body "$WB3" '{"baseRevision":1}'
mreq "$J1" "$CSRF" PUT "/api/projects/$PID" @"$RUN/body.json"
check "stale base → 409" "$CODE" 409
check "409 carries the other version" "$(echo "$BODY" | jq_ workbook)" "$WB2"
check "409 carries its revision" "$(echo "$BODY" | jq_ project.revision)" 2
body "$WB3" '{"name":"x"}'
mreq "$J1" "$CSRF" PUT "/api/projects/$PID" @"$RUN/body.json"
check "save without baseRevision → 422" "$CODE" 422

echo "== versions"
body "$WB3" '{"baseRevision":2}'
mreq "$J1" "$CSRF" PUT "/api/projects/$PID" @"$RUN/body.json"
check "third save 200" "$CODE" 200
req "$J1" GET "/api/projects/$PID/revisions"
check "current revision listed" "$(echo "$BODY" | jq_ current.revision)" 3
check "two earlier versions" "$(echo "$BODY" | php -r '$d=json_decode(stream_get_contents(STDIN),true); echo count($d["revisions"]);')" 2
mreq "$J1" "$CSRF" POST "/api/projects/$PID/revisions/1/restore" '{}'
check "restore 200" "$CODE" 200
check "restore brings back version 1" "$(echo "$BODY" | jq_ workbook)" "$WB1"
check "restore is a new revision" "$(echo "$BODY" | jq_ project.revision)" 4
mreq "$J1" "$CSRF" POST "/api/projects/$PID/revisions/99/restore" '{}'
check "unknown version → 404" "$CODE" 404

echo "== share"
mreq "$J1" "$CSRF" POST "/api/projects/$PID/share" '{"mode":"view"}'
check "share 201" "$CODE" 201
TOKEN=$(echo "$BODY" | jq_ token)
check "share url opens the app" "$(echo "$BODY" | jq_ url)" "$APP/?s=$TOKEN"
req "$RUN/none.txt" GET "/api/share/$TOKEN"
check "anyone can fetch a shared copy" "$CODE" 200
check "shared workbook" "$(echo "$BODY" | jq_ workbook)" "$WB1"
OLDLINK=$(curl -s -o /dev/null -w '%{http_code} %{redirect_url}' "$APP/share/$TOKEN")
check "old /share/ links redirect" "$OLDLINK" "302 $APP/?s=$TOKEN"
mreq "$J1" "$CSRF" DELETE "/api/projects/$PID/share" ''
check "stop sharing 204" "$CODE" 204
req "$RUN/none.txt" GET "/api/share/$TOKEN"
check "revoked link → 404" "$CODE" 404

echo "== workbooks from the first Grid web app"
LEGACY=$(php -r 'echo json_encode(["version"=>1,"title"=>"Old","activeSheet"=>"s1","sheets"=>[["id"=>"s1","name"=>"Sheet1","cells"=>["A1"=>["v"=>42],"B1"=>["v"=>"hello"],"C1"=>["f"=>"A1*2","v"=>84]]]],"styles"=>new stdClass,"names"=>[],"settings"=>new stdClass]);')
LID=11111111-2222-4333-8444-555555555555
php -r '$db=new PDO("sqlite:".$argv[1]); $s=$db->prepare("INSERT INTO grid_projects (id,user_id,name,document,format,revision,bytes,sheets,starred,archived,folder,color,tags,notes,created_at,updated_at) VALUES (?,?,?,?,\"json\",1,?,1,0,0,\"\",\"\",\"[]\",\"\",datetime(\"now\"),datetime(\"now\"))"); $s->execute([$argv[2],"u-hamzeh","Old",$argv[3],strlen($argv[3])]);' "$GRID_SQLITE" "$LID" "$LEGACY"
req "$J1" GET "/api/projects/$LID"
check "legacy get 200" "$CODE" 200
LEG=$(echo "$BODY" | jq_ workbook)
php -r '$b=base64_decode($argv[1],true); $t=tempnam(sys_get_temp_dir(),"l"); file_put_contents($t,$b); $z=new ZipArchive(); exit($z->open($t)===true && $z->locateName("xl/workbook.xml")!==false && str_contains($z->getFromName("xl/worksheets/sheet1.xml"),"A1*2") ? 0 : 1);' "$LEG" \
  && ok "legacy document converted to .xlsx with its formula" || bad "legacy conversion"
body "$WB1" '{"baseRevision":1}'
mreq "$J1" "$CSRF" PUT "/api/projects/$LID" @"$RUN/body.json"
check "legacy workbook saves as .xlsx" "$CODE" 200
php -r '$db=new PDO("sqlite:".$argv[1]); echo $db->query("SELECT format FROM grid_projects WHERE id=\"".$argv[2]."\"")->fetchColumn();' "$GRID_SQLITE" "$LID" > "$RUN/fmt.txt"
check "stored format is now xlsx" "$(cat "$RUN/fmt.txt")" xlsx
mreq "$J1" "$CSRF" POST "/api/projects/$LID/revisions/1/restore" '{}'
check "restoring the legacy version converts it again" "$(echo "$BODY" | jq_ project.revision)" 3

echo "== isolation"
login "$J2" u-eyas >/dev/null
req "$J2" GET /api/me
CSRF2=$(echo "$BODY" | jq_ csrf)
req "$J2" GET "/api/projects/$PID"
check "other user's workbook → 404" "$CODE" 404
body "$WB1" '{"baseRevision":4}'
mreq "$J2" "$CSRF2" PUT "/api/projects/$PID" @"$RUN/body.json"
check "other user can't save over it → 404" "$CODE" 404

echo "== copies, names, delete, limits"
mreq "$J1" "$CSRF" POST "/api/projects/$PID/duplicate" '{}'
check "duplicate 201" "$CODE" 201
DUP=$(echo "$BODY" | jq_ project.id)
req "$J1" GET "/api/projects/$DUP"
check "duplicate has the same workbook" "$(echo "$BODY" | jq_ workbook)" "$WB1"
mreq "$J1" "$CSRF" PATCH "/api/projects/$DUP" '{"name":"Copy  ","starred":true}'
check "rename + star" "$(echo "$BODY" | jq_ project.name)/$(echo "$BODY" | jq_ project.starred)" "Copy/true"
mreq "$J1" "$CSRF" DELETE "/api/projects/$DUP" ''
check "delete 204" "$CODE" 204
req "$J1" GET "/api/projects/$DUP"
check "deleted → 404" "$CODE" 404
N=0
for i in 1 2 3 4 5 6; do
  body "$WB1" "{\"name\":\"n$i\"}"
  mreq "$J2" "$CSRF2" POST /api/projects @"$RUN/body.json"
  [ "$CODE" = 201 ] && N=$((N+1))
done
check "free plan stops at 5 cloud workbooks" "$N/$CODE" "5/429"

echo "== logout and misc"
mreq "$J1" "$CSRF" POST /auth/logout ''
check "logout 204" "$CODE" 204
req "$J1" GET /api/me
check "signed out after logout" "$CODE" 401
req "$RUN/none.txt" GET /api/templates
check "removed routes are gone" "$CODE" 404
req "$RUN/none.txt" DELETE /api/me ''
check "wrong method 405" "$CODE" 405
req "$RUN/none.txt" GET /api/me -D "$RUN/h.txt" >/dev/null
grep -qi 'frame-ancestors' "$RUN/h.txt" && ok "CSP header present" || bad "CSP header"

echo
echo "passed: $PASS   failed: $FAIL"
[ "$FAIL" -eq 0 ]
