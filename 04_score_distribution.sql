SELECT Score_Band, Job_Count, Pct_of_Scored, APPLY_Count, SKIP_Count
FROM (
    SELECT
        CASE
            WHEN Total_Score >= 0.80 THEN '1. 0.80+ High'
            WHEN Total_Score >= 0.75 THEN '2. 0.75-0.79 Good'
            WHEN Total_Score >= 0.70 THEN '3. 0.70-0.74 Solid'
            WHEN Total_Score >= 0.65 THEN '4. 0.65-0.69 Marginal'
            ELSE '5. Below 0.65'
        END AS Score_Band,
        COUNT(*) AS Job_Count,
        ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM `upwork-acquisition-system.acquisition_data.job_scoring` WHERE Total_Score IS NOT NULL), 1) AS Pct_of_Scored,
        SUM(CASE WHEN Final_Decision = 'APPLY' THEN 1 ELSE 0 END) AS APPLY_Count,
        SUM(CASE WHEN Final_Decision = 'SKIP' THEN 1 ELSE 0 END) AS SKIP_Count
    FROM `upwork-acquisition-system.acquisition_data.job_scoring`
    WHERE Total_Score IS NOT NULL
    GROUP BY Score_Band

    UNION ALL

    SELECT
        'AVG_SCORE',
        NULL,
        NULL,
        ROUND(AVG(Total_Score), 2),
        NULL
    FROM `upwork-acquisition-system.acquisition_data.job_scoring`
    WHERE Total_Score IS NOT NULL
)
ORDER BY Score_Band;
