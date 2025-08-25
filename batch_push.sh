#!/bin/bash
# スクリプトIDのリスト
script_ids=(
"1jYgaYX3LnPNmdcppC7NgLAQS2idOkO3vr5h6MrO3NddBfWfQCQrI9b_f"
"1DTjL6HArs2ypoK6nQ1a0MulOlrsf-L3HPQrnOtsnlSN_3gZju95Ndn-a"
"1uBCd3fuSmOJ_bkrmkxUFz9GNYC9FeGPiQLCjcKQz_-MHffeGoZHDZ9ie"
"1EWDmtmC7y20n4zdHtIeBD03OrTYlOxbrjGi-FOy-T-f3DQBF_8rzfkzj"
"1SjXdlzCFrIolSAyYBvvrgghfJyuVyGcyOyprF61YlniV4nrCZkJjFZY9"
"1pOTZTxtPeIJv6U_bAZJT6H30ifXJWIRrwsKBkDHuELG5EvdHs9Es5Ex2"
"1T1QGFmR7rzgjzQTQ52flk3vOOIte0uxoeYGliyacigw1hMZY4D678hkr"
"13kfqCyldp9NNlnId2t-eKBcrRYTqQvVU25fWyZYPbZ_0l79eBwTgAz3t"
)

#!/bin/bash
# バッチプッシュスクリプト - セキュリティ警告付き

echo "⚠️  警告: このスクリプトは8つのプロジェクトに対して --force フラグでプッシュを実行します"
echo "   これは危険な操作です。すべてのプロジェクトが上書きされます。"
echo ""
read -p "本当に続行しますか？ (yes/NO): " confirm

if [[ "$confirm" != "yes" ]]; then
    echo "キャンセルされました。"
    exit 1
fi

echo ""
echo "🔄 プッシュを開始します..."
echo ""

# 現在の.clasp.jsonをバックアップ
cp .clasp.json .clasp.json.backup

# 各スクリプトIDに対してプッシュ
for i in "${!script_ids[@]}"; do
  script_id="${script_ids[$i]}"
  echo "[$((i+1))/${#script_ids[@]}] プッシュ中: $script_id"
  
  # .clasp.jsonを更新
  echo "{\"scriptId\":\"$script_id\",\"rootDir\":\"src\"}" > .clasp.json
  
  # プッシュ実行
  clasp push --force
  
  if [ $? -eq 0 ]; then
    echo "✓ 成功: $script_id"
  else
    echo "✗ 失敗: $script_id"
  fi
  
  echo "---"
done

# .clasp.jsonを元に戻す
mv .clasp.json.backup .clasp.json

echo "すべてのプッシュが完了しました。"