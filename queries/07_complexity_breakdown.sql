SELECT
    TRIM(SPLIT(Quick_Notes, ',')[OFFSET(0)]) AS Complexity,
    COUNT(*) AS Job_Count,
    ROUND(COUNT(*) * 100.0 /
        (SELECT COUNT(*) 
         FROM `upwork-acquisition-system.acquisition_data.job_discovery` 
         WHERE Quick_Notes IS NOT NULL AND Quick_Notes != ''), 1) AS Pct_of_Total
FROM `upwork-acquisition-system.acquisition_data.job_discovery`
WHERE Quick_Notes IS NOT NULL AND Quick_Notes != ''
GROUP BY Complexity
ORDER BY Job_Count DESC;
