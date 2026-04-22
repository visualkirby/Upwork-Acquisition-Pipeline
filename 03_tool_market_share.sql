SELECT
    Tool_Detected,
    COUNT(*) AS Total_Jobs,
    ROUND(COUNT(*) * 100.0 /
        (SELECT COUNT(*) 
         FROM `upwork-acquisition-system.acquisition_data.job_discovery` 
         WHERE Tool_Detected != ''), 1) AS Market_Share_Pct,
    SUM(CASE WHEN Discovery_Action = 'Move to Scoring' 
        THEN 1 ELSE 0 END) AS Moved_to_Scoring,
    ROUND(SUM(CASE WHEN Discovery_Action = 'Move to Scoring' 
        THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS Scoring_Rate_Pct
FROM `upwork-acquisition-system.acquisition_data.job_discovery`
WHERE Tool_Detected != ''
GROUP BY Tool_Detected
ORDER BY Total_Jobs DESC;
