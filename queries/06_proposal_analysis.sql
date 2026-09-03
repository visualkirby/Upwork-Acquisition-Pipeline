SELECT
    Template_Used, Hook_Version, CTA_Version,
    COUNT(*) AS Proposals_Sent,
    ROUND(AVG(CAST(Connects_Used AS FLOAT64)), 1) AS Avg_Connects,
    SUM(CASE WHEN Client_Replied = TRUE THEN 1 ELSE 0 END) AS Replies,
    SUM(CASE WHEN Interview = TRUE THEN 1 ELSE 0 END) AS Interviews,
    SUM(CASE WHEN Hired = TRUE THEN 1 ELSE 0 END) AS Hires,
    ROUND(SUM(CASE WHEN Client_Replied = TRUE THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS Reply_Rate_Pct,
    ROUND(SUM(CASE WHEN Interview = TRUE THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS Interview_Rate_Pct,
    ROUND(SUM(CASE WHEN Hired = TRUE THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS Hire_Rate_Pct
FROM `upwork-acquisition-system.acquisition_data.proposal_tracker`
GROUP BY Template_Used, Hook_Version, CTA_Version

UNION ALL

SELECT
    'OVERALL' AS Template_Used,
    NULL AS Hook_Version,
    NULL AS CTA_Version,
    COUNT(*) AS Proposals_Sent,
    ROUND(AVG(CAST(Connects_Used AS FLOAT64)), 1) AS Avg_Connects,
    SUM(CASE WHEN Client_Replied = TRUE THEN 1 ELSE 0 END) AS Replies,
    SUM(CASE WHEN Interview = TRUE THEN 1 ELSE 0 END) AS Interviews,
    SUM(CASE WHEN Hired = TRUE THEN 1 ELSE 0 END) AS Hires,
    ROUND(SUM(CASE WHEN Client_Replied = TRUE THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS Reply_Rate_Pct,
    ROUND(SUM(CASE WHEN Interview = TRUE THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS Interview_Rate_Pct,
    ROUND(SUM(CASE WHEN Hired = TRUE THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) AS Hire_Rate_Pct
FROM `upwork-acquisition-system.acquisition_data.proposal_tracker`

ORDER BY Proposals_Sent DESC;
