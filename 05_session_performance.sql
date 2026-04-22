SELECT
    Session_ID,
    Date,
    Duration,
    Jobs_Logged,
    Jobs_Moved_To_Scoring,
    Jobs_Review_Later,
    Session_Yield,
    Saturation_Flag,
    COALESCE(Proposals_Sent, 0) AS Proposals_Sent,
    COALESCE(Connects_Spent, 0) AS Connects_Spent
FROM `upwork-acquisition-system.acquisition_data.session_log`
ORDER BY Session_ID;
