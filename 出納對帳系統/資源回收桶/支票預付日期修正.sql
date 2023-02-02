update 傳票資料 set 預付日期 = 登錄日期, 登錄日期 = NULL
FROM 傳票資料 INNER JOIN 現金備查簿 ON 傳票資料.年 = 現金備查簿.年 AND 傳票資料.傳票號碼 = 現金備查簿.傳票號碼 
WHERE (現金備查簿.收入金額405 > 0 OR 現金備查簿.支出金額405 > 0) 
AND NOT (傳票資料.匯入帳號 IS NOT NULL AND 傳票資料.匯入帳號 != '')
AND (傳票資料.收入金額 = 0 OR 傳票資料.收入金額 IS NULL)
