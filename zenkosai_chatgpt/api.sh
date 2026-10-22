#!/bin/bash
# usage: api.sh key=value ...   (values URL-encoded automatically)
URL="${ZENKOSAI_URL:?ZENKOSAI_URL（ChatGPT設定シートB1のURL）を設定してください}"
args=(--data-urlencode "token=${ZENKOSAI_TOKEN:?ZENKOSAI_TOKEN（ルールシートの合言葉）を設定してください}")
for kv in "$@"; do args+=(--data-urlencode "$kv"); done
curl -sS -L -G --max-time 90 "${args[@]}" "$URL"; echo
