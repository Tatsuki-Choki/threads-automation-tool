#!/bin/bash

# 残りのスクリプトIDのリスト（18番目から）
SCRIPT_IDS=(
  "1jYgaYX3LnPNmdcppC7NgLAQS2idOkO3vr5h6MrO3NddBfWfQCQrI9b_f"
  "1DTjL6HArs2ypoK6nQ1a0MulOlrsf-L3HPQrnOtsnlSN_3gZju95Ndn-a"
  "1uBCd3fuSmOJ_bkrmkxUFz9GNYC9FeGPiQLCjcKQz_-MHffeGoZHDZ9ie"
  "1EWDmtmC7y20n4zdHtIeBD03OrTYlOxbrjGi-FOy-T-f3DQBF_8rzfkzj"
  "1-AcFjmLI0FFkFtp-Qlbyl3jtOSilFWUDox7mjcoXbl60EmzBB5GOsOnI"
  "1SjXdlzCFrIolSAyYBvvrgghfJyuVyGcyOyprF61YlniV4nrCZkJjFZY9"
  "1pOTZTxtPeIJv6U_bAZJT6H30ifXJWIRrwsKBkDHuELG5EvdHs9Es5Ex2"
  "1T1QGFmR7rzgjzQTQ52flk3vOOIte0uxoeYGliyacigw1hMZY4D678hkr"
  "13kfqCyldp9NNlnId2t-eKBcrRYTqQvVU25fWyZYPbZ_0l79eBwTgAz3t"
  "1kqblqSzC9rZnO0CjOb6Vy17EI9ae0q4UBnDhDrn_zZDWxijZczmWNadq"
)

# 現在の.clasp.jsonをバックアップ
if [ -f ".clasp.json" ]; then
  cp .clasp.json .clasp.json.backup
  echo "既存の.clasp.jsonをバックアップしました"
fi

# 成功・失敗カウンター
SUCCESS_COUNT=0
FAIL_COUNT=0
FAILED_IDS=()

echo "===================="
echo "残りのスクリプトへのプッシュを開始します"
echo "合計: ${#SCRIPT_IDS[@]} 個のスクリプト（18-27）"
echo "===================="
echo ""

# 各スクリプトIDに対してプッシュ
START_NUM=18
for i in "${!SCRIPT_IDS[@]}"; do
  SCRIPT_ID="${SCRIPT_IDS[$i]}"
  CURRENT_NUM=$((START_NUM + i))
  echo "[$CURRENT_NUM/27] プッシュ中: $SCRIPT_ID"
  
  # .clasp.jsonを更新
  cat > .clasp.json << EOF
{
  "scriptId": "$SCRIPT_ID",
  "rootDir": "src"
}
EOF
  
  # clasp pushを実行
  if clasp push 2>&1 | grep -q "Pushed"; then
    echo "✅ 成功"
    ((SUCCESS_COUNT++))
  else
    echo "❌ 失敗"
    ((FAIL_COUNT++))
    FAILED_IDS+=("$SCRIPT_ID")
  fi
  
  echo "--------------------"
  
  # 短い待機時間（APIレート制限対策）
  sleep 1
done

echo ""
echo "===================="
echo "プッシュ完了"
echo "===================="
echo "成功: $SUCCESS_COUNT 個"
echo "失敗: $FAIL_COUNT 個"

if [ ${#FAILED_IDS[@]} -gt 0 ]; then
  echo ""
  echo "失敗したスクリプトID:"
  for ID in "${FAILED_IDS[@]}"; do
    echo "  - $ID"
  done
fi

# .clasp.jsonを元に戻す
if [ -f ".clasp.json.backup" ]; then
  mv .clasp.json.backup .clasp.json
  echo ""
  echo ".clasp.jsonを元に戻しました"
fi

echo ""
echo "完了しました！"