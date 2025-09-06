#!/usr/bin/env bash
set -e

# Usage:
#   bash push_ids.sh <SCRIPT_ID_1> <SCRIPT_ID_2> ...
#   # または、script_ids.txt（非空行のみ）にIDを列挙して実行:
#   bash push_ids.sh

: > push_results.txt

if [ $# -eq 0 ] && [ -f script_ids.txt ]; then
  # 空行を除外して引数化
  mapfile -t ARGS < <(grep -E '\S' script_ids.txt)
  set -- "${ARGS[@]}"
fi

for id in "$@"; do
  echo "--- Pushing $id"
  printf '{\n  "scriptId": "%s",\n  "rootDir": "src"\n}\n' "$id" > .clasp.json
  if npx --yes clasp push -f; then
    echo "$id: OK" >> push_results.txt
  else
    echo "$id: FAIL" >> push_results.txt
  fi
done

echo "Summary:"
cat push_results.txt

#!/usr/bin/env bash
set -e
: > push_results.txt
for id in "$@"; do
  echo "--- Pushing $id"
  printf {n