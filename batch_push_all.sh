#!/bin/bash

# 全スクリプトIDのリスト
SCRIPT_IDS=(
  "1iXaD49YSZePJqRhN1_cKfmW4dNAZYQzWEOEItp2Zvl6SWOF9QMMvSJhV"
  "10HfOF5g9eu9WpXfaY103Mp8AsMDNjbqVgiopEbpWchvu8oswgRMypLIt"
  "18Rd4kWm2br12UaPjXErowzd0P53uT4S5AJOr_nvxMZ4w1-fpANSFnaNB"
  "16y3Lg47w0gpvBqGq2hksZ3jcCl7gVlNqHgDmHu5VqBZSUFLkZ_EtQkou"
  "1FKQ-5fzEwwJtlc-lvl5SDuK9sZbE0pa-UcimOkBXTEsWBxE_TkZrpM94"
  "1iwjT9e9smH_6LwrDGd-eu5O17V-I6I_EAkRifPPmXemP_cAR7soolMrK"
  "1PsfBI2tzXoj-WVJcRwwxBzBPDfTrDvfH7-eZzSBnn0SoJyXWCCrlD__2"
  "1HX25vALFetweKe-tjQe_th4Jdx4w-1G2WuOMfQOaQ7RimB91RBIYYbIw"
  "1qzI_oo5s8wla3oyb4TvAQUTiis__kob-ux9O6ocvhUpENE2YNZDHNvKZ"
  "19-z0XzMCKC1w_TNUBUq21aRNKwoiauGIPNYdkefgM_7OZMeJZvuAIbhC"
  "1JQm8wmcnGcr43YKHH7v4OiofI-m1DSQoXV36PYzhnTFtBAcQOQI3Et7M"
  "1IG2NmC3QEgA0mRpJBvBUKc8RBdfM9JUdOKraJkZ9v6Ok5zG90TIzvbtB"
  "1W9e6jdJihjmWd13sccuN303ALuiQROfzf_eh07lYrQ2Si_jxfx8Mtfpr"
  "15EmpEc5XNcMCH5eFPdHzMmk2EiWd76nF2WecZPYan9A3EUou02QY3JZ3"
  "1NF-J8X51De3je4WFX482JAECF1b4baxaNn7y6ie9AGeI1fytYqJLLp8b"
  "1if_xTTPYHxredMaS4zKo1nLXnR4R96rZflaXG7shtEqPOTMGVRKToTTV"
  "1WqtnJ-k24sDO5TtZANhHvsI5tvV4fDhfUmtBPzxFnG0-C2f9mqf5df4v"
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
echo "全スクリプトへのプッシュを開始します"
echo "合計: ${#SCRIPT_IDS[@]} 個のスクリプト"
echo "===================="
echo ""

# 各スクリプトIDに対してプッシュ
for i in "${!SCRIPT_IDS[@]}"; do
  SCRIPT_ID="${SCRIPT_IDS[$i]}"
  echo "[$((i+1))/${#SCRIPT_IDS[@]}] プッシュ中: $SCRIPT_ID"
  
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