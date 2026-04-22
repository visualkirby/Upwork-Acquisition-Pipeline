SELECT
    'Discovered'    AS Stage,
    1               AS Stage_Order,
    COUNT(*)        AS Job_Count,
    ROUND(COUNT(*) * 100.0 / COUNT(*), 1) AS Pct_of_Discovered
FROM `upwork-acquisition-system.acquisition_data.job_discovery`
WHERE Discovery_Action != ''

UNION ALL

SELECT
    'Moved to Scoring',
    2,
    COUNT(*),
    ROUND(COUNT(*) * 100.0 /
        (SELECT COUNT(*) FROM `upwork-acquisition-system.acquisition_data.job_discovery`
         WHERE Discovery_Action != ''), 1)
FROM `upwork-acquisition-system.acquisition_data.job_discovery`
WHERE Discovery_Action = 'Move to Scoring'

UNION ALL

SELECT
    'APPLY Decision',
    3,
    COUNT(*),
    ROUND(COUNT(*) * 100.0 /
        (SELECT COUNT(*) FROM `upwork-acquisition-system.acquisition_data.job_discovery`
         WHERE Discovery_Action != ''), 1)
FROM `upwork-acquisition-system.acquisition_data.job_scoring`
WHERE Final_Decision = 'APPLY'

UNION ALL

SELECT
    'Proposal Sent',
    4,
    COUNT(*),
    ROUND(COUNT(*) * 100.0 /
        (SELECT COUNT(*) FROM `upwork-acquisition-system.acquisition_data.job_discovery`
         WHERE Discovery_Action != ''), 1)
FROM `upwork-acquisition-system.acquisition_data.proposal_tracker`

ORDER BY Stage_Order;
