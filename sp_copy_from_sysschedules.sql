CREATE PROCEDURE [dbo].[sp_copy_from_sysschedules]
(
    @sysschedule_id INT,
    @schedule_id INT OUTPUT
)
AS
BEGIN
    INSERT INTO [dbo].[schedules]
    (
        [name],
        [freq_type],
        [freq_interval],
        [freq_subday_type],
        [freq_subday_interval],
        [freq_relative_interval],
        [freq_recurrence_factor],
        [active_start_date],
        [active_end_date],
        [active_start_time],
        [active_end_time]
    )
    
    SELECT
        [name],
        [freq_type],
        [freq_interval],
        [freq_subday_type],
        [freq_subday_interval],
        [freq_relative_interval],
        [freq_recurrence_factor],
        [active_start_date],
        [active_end_date],
        [active_start_time],
        [active_end_time]
    FROM [msdb].[dbo].[sysschedules]
    WHERE [schedule_id] = @sysschedule_id;
    
    SET @schedule_id = SCOPE_IDENTITY();
END;