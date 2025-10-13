#!/bin/bash

# 修正済みコードを全スプレッドシートにプッシュするスクリプト

SCRIPT_IDS=(
"1MenZk9WxpZAv0U79xESYwz3aLENP_toLQuanvx_f7imC5uPijTOUBj5k"
"10HfOF5g9eu9WpXfaY103Mp8AsMDNjbqVgiopEbpWchvu8oswgRMypLIt"
"18Rd4kWm2br12UaPjXErowzd0P53uT4S5AJOr_nvxMZ4w1-fpANSFnaNB"
"16y3Lg47w0gpvBqGq2hksZ3jcCl7gVlNqHgDmHu5VqBZSUFLkZ_EtQkou"
"1FKQ-5fzEwwJtlc-lvl5SDuK9sZbE0pa-UcimOkBXTEsWBxE_TkZrpM94"
"1iwjT9e9smH_6LwrDGd-eu5O17V-I6I_EAkRifPPmXemP_cAR7soolMrK"
"1PsfBI2tzXoj-WVJcRwwxBzBPDfTrDvfH7-eZzSBnn0SoJyXWCCrlD__2"
"1HX25vALFetweKe-tjQe_th4Jdx4w-1G2WuOMfQOaQ7RimB91RBIYYbIw"
"1qzI_oo5s8wla3oyb4TvAQUTiis__kob-ux9O6ocvhUpENE2YNZDHNvKZ"
"19-z0XzMCKC1w_TNUBUUq21aRNKwoiauGIPNYdkefgM_7OZMeJZvuAIbhC"
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
"1gmdNcxNWpcN8vvRHy7PuQvijs3dsAeMuuSMl0sLKn-HaigV414LKuzPD"
"1svrzJpqQlIdQuPZstaUyf_ssWoqsM3AUV0lrPt8fMDhN8oQSlB5tOaIN"
"17Oz6-pIydWbAz3ExAlTUI_a1fAyqfnYi2Si-DOXuYgrP4i73U0NdnMgW"
"1-kQbO7n11jhmbW-RQPLJfI0D1frLAq-YemYWZZzBdjtdfO1txRSs4MYm"
"1lfH8NrpCwkzlIF6kcXgfPj-gwuG4oRFjpAK7U07XlOH3ymIDPyCIOtKh"
"1910E5bBkGoKc8PBvgiAPDzbqceXwrychRC3NqYh6SnpgU3oVa-0PMEPl"
"1-AcFjmLI0FFkFtp-Qlbyl3jtOSilFWUDox7mjcoXbl60EmzBB5GOsOnI"
"1SjXdlzCFrIolSAyYBvvrgghfJyuVyGcyOyprF61YlniV4nrCZkJjFZY9"
"1pOTZTxtPeIJv6U_bAZJT6H30ifXJWIRrwsKBkDHuELG5EvdHs9Es5Ex2"
"1T1QGFmR7rzgjzQTQ52flk3vOOIte0uxoeYGliyacigw1hMZY4D678hkr"
"13kfqCyldp9NNlnId2t-eKBcrRYTqQvVU25fWyZYPbZ_0l79eBwTgAz3t"
"1kqblqSzC9rZnO0CjOb6Vy17EI9ae0q4UBnDhDrn_zZDWxijZczmWNadq"
"1gzdjdvzJFJWAgSI54eR2atFIyBvawmrTfWCfPbDFym1mEO2YeOIvJP-X"
"19qINnBKLYZJkKToo355e-z2x1sdF18h6qITx2IWlX4mdfCSsHTBu_OoF"
"1HFLrjxgjEtHuUTZvHDAQzif1USfBUCnfy8vRHsedH9c9hYB8CBXaHmKn"
"1JkgQ9mgiPZmZUvPN31V83fkYoYT7VawxdIqWg_JpZzl4oCyhdROWt8-i"
"15cvTzlc-4RKFriZ8mu6ny31U4guTb0EAGn9VnB2FrVsNW0vi5-dBxZvX"
"18nJS1eRWN0DMlk1tmHz_UAdvqvSWgVNfPO9o-V-STnyiqlMECIuOQtYi"
"12FG0vlHGDIOyAFTzMjQNJUrEvE_7U5LHFlt-QnvS7Kau3R3UWb-fcOVJ"
"1Lh8DswJuuBCNYFku_CFEsR5OvPxNHKl1Xbj6KS-webUtSJJq9VyKkxiU"
"1XXCJSbiNhWfOXsljVm8CnJ0_XLlvtQuHhUJ96pgyq4Zop5c3gyrMd8Li"
"19Ic37rsZTxnbB4GRlPQ0NeHNmzwBbACWcPZefbQ0ZvxL1pxxBsSWHh7E"
"1G1V05HPxuNwQZUNxd9ui4CPXHFT2oamPTZcpJGc8hSQfyr83vJrkFfjx"
"1z0MnMoUbbfCTsADDYzUOkKlLT1vYx4Z5pm_gEeAjxMYp3iM8X7Gt1R2L"
"1lL9NucJ-0G8Yb2qZXisq8KPsVc0u3zVv8jjp2LoLfZmljGmDB3ZMK1Tz"
"1G-u7AD-lsLbd0d5am0H_px_GqQ-yZJAeSjbEaZ26cwG_v_j20EC5JtCU"
"1OebtS8KzeQLO7ADhhU8KJXcOtvaj9lSKrBwELzLHI9i5Fz5jSea8y0eU"
"1kPWseZH-_3oLMUI3EnmTTzSgf0F2Bawpj_ImJBk7nhP3YSrLXwTn01cn"
"1OXy1M0IQGjIeiELCEheMFYhZ1XAGthK2bcX1KIkneFRC9CsLH6wo866e"
"14PMaBiUIXIZjUibyofSS9YVBTjxcXtaC4fZ5Kg4Awf2FRx4YHjyQ1vP6"
"1SaviNSprZkCdPLP4U6MI_W6qbma5F5uWab1iy08v86k2_ySf4c9HXemk"
"1e49eaCzBb9QiKCzbawWof3k6TP1-3eM2IggJACIKEe-XXsJlEqV_j_q6"
"1q3PmYFjaFUF0UTNTp6B1Nh7ovluEux4aKJ3pctvHiaukjkt4B270j878"
"15LNyMsayjVIBtFp8Ez6I_4gqatYHdzIxT2hikCpWKOR1snEKrOSRKZiT"
"1miXvK6QZ5RIcnpmhGVp56YIU4UdxLP9q_Gsp9GdzOpqDBpzA1cPV0359"
"1Nrx3ePtKg1CdBkscc99AwsGUpncjT6gsGnpelYsIjvHeeqqTMEvTCFQW"
"1vxynNcyEi2jgpRNeVN6iQVAKWawJ-KtthWA5CNiNKxuD9GPmdM32RycY"
"1BfKRIB2WakRUBQGoCPDwQgwtyUGYc35h8z3m7LvkMdHuXrMTLWOpJ-cW"
)

TOTAL=${#SCRIPT_IDS[@]}
SUCCESS=0
FAILED=0
FAILED_IDS=()

echo "========================================="
echo "修正済みコードの一括プッシュ"
echo "========================================="
echo "対象スクリプト数: $TOTAL"
echo ""

for i in "${!SCRIPT_IDS[@]}"; do
    SCRIPT_ID="${SCRIPT_IDS[$i]}"
    INDEX=$((i + 1))
    
    echo "[$INDEX/$TOTAL] プッシュ中: $SCRIPT_ID"
    
    # .clasp.jsonを更新
    echo "{\"scriptId\": \"$SCRIPT_ID\"}" > .clasp.json
    
    # プッシュ実行
    if clasp push --force > /dev/null 2>&1; then
        echo "  ✓ 成功"
        SUCCESS=$((SUCCESS + 1))
    else
        echo "  ✗ 失敗"
        FAILED=$((FAILED + 1))
        FAILED_IDS+=("$SCRIPT_ID")
    fi
    
    echo ""
done

echo "========================================="
echo "プッシュ完了"
echo "========================================="
echo "成功: $SUCCESS / $TOTAL"
echo "失敗: $FAILED / $TOTAL"
echo ""

if [ $FAILED -gt 0 ]; then
    echo "失敗したスクリプトID:"
    for ID in "${FAILED_IDS[@]}"; do
        echo "  - $ID"
    done
    echo ""
fi

echo "完了"

