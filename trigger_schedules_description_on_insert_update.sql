CREATE TRIGGER [dbo].[trigger_schedules_description_on_insert_update]
ON [dbo].[schedules]
AFTER INSERT, UPDATE
AS
BEGIN
    DECLARE @desc NVARCHAR(255);
    DECLARE @freq_type INT, @freq_interval INT, @freq_subday_type INT, @freq_subday_interval INT, @freq_relative_interval INT, @freq_recurrence_factor INT, @active_start_date INT, @active_end_date INT, @active_start_time INT, @active_end_time INT, @schedule_id INT;

    DECLARE inserted_cursor CURSOR LOCAL FOR
    SELECT freq_type, freq_interval, freq_subday_type, freq_subday_interval, freq_relative_interval, freq_recurrence_factor, active_start_date, active_end_date, active_start_time, active_end_time, schedule_id
    FROM inserted;

    OPEN inserted_cursor;
    FETCH NEXT FROM inserted_cursor INTO @freq_type, @freq_interval, @freq_subday_type, @freq_subday_interval, @freq_relative_interval, @freq_recurrence_factor, @active_start_date, @active_end_date, @active_start_time, @active_end_time, @schedule_id;

    WHILE @@FETCH_STATUS = 0
    BEGIN
        EXEC [dbo].[sp_get_schedule_description] 
            @freq_type = @freq_type,
            @freq_interval = @freq_interval,
            @freq_subday_type = @freq_subday_type,
            @freq_subday_interval = @freq_subday_interval,
            @freq_relative_interval = @freq_relative_interval,
            @freq_recurrence_factor = @freq_recurrence_factor,
            @active_start_date = @active_start_date,
            @active_end_date = @active_end_date,
            @active_start_time = @active_start_time,
            @active_end_time = @active_end_time,
            @schedule_description = @desc output;

        UPDATE [dbo].[schedules]
        SET description = @desc
        WHERE [dbo].[schedules].schedule_id = @schedule_id;

        FETCH NEXT FROM inserted_cursor INTO @freq_type, @freq_interval, @freq_subday_type, @freq_subday_interval, @freq_relative_interval, @freq_recurrence_factor, @active_start_date, @active_end_date, @active_start_time, @active_end_time, @schedule_id;
    END;

    CLOSE inserted_cursor;
    DEALLOCATE inserted_cursor;
END;