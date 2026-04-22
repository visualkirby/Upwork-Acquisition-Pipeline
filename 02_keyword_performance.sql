SELECT
    js.Keyword_Search,
    ki.Tool_Focus,
    ki.Priority_Score,
    ki.Status,
    COUNT(*) AS Total_Scored,
    SUM(CASE WHEN js.Final_Decision = 'APPLY' THEN 1 ELSE 0 END) AS APPLY_Count,
    SUM(CASE WHEN js.Final_Decision = 'SKIP' THEN 1 ELSE 0 END) AS SKIP_Count,
    SUM(CASE WHEN js.Final_Decision = 'HOLD' THEN 1 ELSE 0 END) AS HOLD_Count,
    ROUND(SUM(CASE WHEN js.Final_Decision = 'APPLY' THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS APPLY_Rate_Pct,
    ROUND(AVG(js.Total_Score), 3) AS Avg_Score,
    ROUND(AVG(CAST(js.Proposal_Count AS FLOAT64)), 1) AS Avg_Competition,
    ki.Connect_Efficiency
FROM `upwork-acquisition-system.acquisition_data.job_scoring` js
INNER JOIN `upwork-acquisition-system.acquisition_data.keyword_intelligence` ki
    ON LOWER(TRIM(js.Keyword_Search)) = LOWER(TRIM(ki.Keyword))
WHERE js.Keyword_Search != ''
GROUP BY js.Keyword_Search, ki.Tool_Focus, ki.Priority_Score, ki.Status, ki.Connect_Efficiency
ORDER BY Priority_Score DESC, Avg_Score DESC;
