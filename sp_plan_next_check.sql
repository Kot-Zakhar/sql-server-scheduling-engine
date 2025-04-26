CREATE PROCEDURE [dbo].[sp_plan_next_check] (@schedule_id INT)
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE
         @freq_type INT
        ,@freq_interval INT
        ,@freq_subday_type INT
        ,@freq_subday_interval INT
        ,@freq_relative_interval INT
        ,@freq_recurrence_factor INT
        ,@active_start_date INT
        ,@active_end_date INT
        ,@active_start_time INT
        ,@active_end_time INT
        ,@next_execution_date INT
        ,@next_execution_time INT
        ,@last_execution_date INT
        ,@last_execution_time INT;

    SELECT TOP 1 @last_execution_date = previous_execution_date,
                @last_execution_time = previous_execution_time
    FROM [dbo].[next_checks]
    WHERE schedule_id = @schedule_id
    ORDER BY next_execution_date DESC, next_execution_time DESC;

    SELECT @freq_type = freq_type,
        @freq_interval = freq_interval,
        @freq_subday_type = freq_subday_type,
        @freq_subday_interval = freq_subday_interval,
        @freq_relative_interval = freq_relative_interval,
        @freq_recurrence_factor = freq_recurrence_factor,
        @active_start_date = active_start_date,
        @active_end_date = active_end_date,
        @active_start_time = active_start_time,
        @active_end_time = active_end_time
    FROM [dbo].[schedules]
    WHERE schedule_id = @schedule_id;

    EXEC [dbo].[sp_get_schedule_next_execution_date_and_time]
        @freq_type
        ,@freq_interval
        ,@freq_subday_type
        ,@freq_subday_interval
        ,@freq_relative_interval
        ,@freq_recurrence_factor
        ,@active_start_date
        ,@active_end_date
        ,@active_start_time
        ,@active_end_time
        ,@last_execution_date
        ,@last_execution_time
        ,@next_execution_date OUTPUT
        ,@next_execution_time OUTPUT;
    
    UPDATE [dbo].[next_checks]
    SET
        next_execution_date = @next_execution_date,
        next_execution_time = @next_execution_time
    WHERE schedule_id = @schedule_id;

    IF @@ROWCOUNT = 0
        INSERT INTO [dbo].[next_checks] (
            schedule_id,
            previous_execution_date,
            previous_execution_time,
            next_execution_date,
            next_execution_time)
        VALUES (
            @schedule_id,
            @last_execution_date,
            @last_execution_time,
            @next_execution_date,
            @next_execution_time);

    -- return the value
    SELECT
        @schedule_id schedule_id,
        @last_execution_date previous_execution_date,
        @last_execution_time previous_execution_time,
        @next_execution_date next_execution_date,
        @next_execution_time next_execution_time;
END