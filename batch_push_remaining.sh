#!/bin/bash

# 失敗したスクリプトIDのみ
script_ids=(
"1_Dq1GZcpkY-Pu4vPXH6LM5Faho7eqr9u5QMH_bRe7dbdIZkFquttqxHt"
"1e49eaCzBb9QiKCzbawWof3k6TP1-3eM2IggJACIKEe-XXsJlEqV_j_q6"
)

echo "===================="
echo "失敗分の再試行"
echo "対象: ${#script_ids[@]}個のスクリプト"
echo "===================="

for id in "${script_ids[@]}"; do
    echo ""
    echo "プッシュ中: $id"
    
    # .clasp.jsonを更新
    echo "{\"scriptId\":\"$id\",\"rootDir\":\"src\"}" > .clasp.json
    
    # clasp push実行（詳細表示）
    clasp push
done

# 元の.clasp.jsonを復元
echo '{"scriptId":"1_Dq1GZcpkY-Pu4vPXH6LM5Faho7eqr9u5QMH_bRe7dbdIZkFquttqxHt","rootDir":"src"}' > .clasp.json
