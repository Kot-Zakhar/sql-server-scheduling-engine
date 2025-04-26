CREATE PROCEDURE [dbo].[sp_get_schedule_description]
(
    @freq_type INT,
    @freq_interval INT,
    @freq_subday_type INT,
    @freq_subday_interval INT,
    @freq_relative_interval INT,
    @freq_recurrence_factor INT,
    @active_start_date INT,
    @active_end_date INT,
    @active_start_time INT,
    @active_end_time INT,
    @schedule_description VARCHAR(MAX) OUTPUT
)
AS
BEGIN
    DECLARE @time_format VARCHAR(13) = 'hh\:mm\:ss tt';
    DECLARE @date_format VARCHAR(10) = 'yyyy-MM-dd';
    DECLARE
        @active_start_date_DATE DATE = TRY_CONVERT(DATE, CONVERT(VARCHAR(8), @active_start_date)),
        @active_start_time_TIME TIME = TRY_CONVERT(TIME, STUFF(STUFF(RIGHT('000000' + CONVERT(VARCHAR(6), @active_start_time), 6), 5, 0, ':'), 3, 0, ':')),
        @active_end_date_DATE DATE = TRY_CONVERT(DATETIME, CONVERT(VARCHAR(8), @active_end_date)),
        @active_end_time_TIME TIME = TRY_CONVERT(TIME, STUFF(STUFF(RIGHT('000000' + CONVERT(VARCHAR(6), @active_end_time), 6), 5, 0, ':'), 3, 0, ':'));

	-- Validation
	IF (@active_end_date <> 0 AND @active_end_date < @active_start_date)
    BEGIN
        SET @schedule_description = 'Error: active date range is not valid.';
		RETURN;
    END;

    IF (@active_end_time <> 0 AND @active_end_time < @active_start_time)
    BEGIN
        SET @schedule_description = 'Error: active time range is not valid.';
        RETURN;
    END;

    IF @freq_type NOT IN (1, 4, 8, 16, 32)
    BEGIN
        SET @schedule_description = 'Error: freq_type ' + CONVERT(VARCHAR, @freq_type) + ' is not supported.';
        RETURN;
    END;

    SET @schedule_description = CASE @freq_type
        WHEN 1 THEN 'Occurs once '
        WHEN 4 THEN 'Occurs every ' + CASE WHEN @freq_interval = 1 THEN 'day ' ELSE CONVERT(VARCHAR, @freq_interval) + ' days ' END
        WHEN 8 THEN 'Occurs ' + 
            CASE @freq_recurrence_factor
                WHEN 0 THEN 'weekly '
                WHEN 1 THEN 'every week '
                ELSE 'every ' + CONVERT(VARCHAR, @freq_recurrence_factor) + ' weeks '
            END +
            'on ' +
            CASE WHEN (@freq_interval & 1) > 0 THEN ('Sunday' + CASE WHEN (@freq_interval & (64 + 32 + 16 + 8 + 4 + 2)) > 0 THEN ', ' ELSE '' END) ELSE '' END + 
            CASE WHEN (@freq_interval & 2) > 0 THEN ('Monday' + CASE WHEN (@freq_interval & (64 + 32 + 16 + 8 + 4)) > 0 THEN ', ' ELSE '' END) ELSE '' END + 
            CASE WHEN (@freq_interval & 4) > 0 THEN ('Tuesday' + CASE WHEN (@freq_interval & (64 + 32 + 16 + 8)) > 0 THEN ', ' ELSE '' END) ELSE '' END + 
            CASE WHEN (@freq_interval & 8) > 0 THEN ('Wednesday' + CASE WHEN (@freq_interval & (64 + 32 + 16)) > 0 THEN ', ' ELSE '' END) ELSE '' END + 
            CASE WHEN (@freq_interval & 16) > 0 THEN ('Thursday' + CASE WHEN (@freq_interval & (64 + 32)) > 0 THEN ', ' ELSE '' END) ELSE '' END + 
            CASE WHEN (@freq_interval & 32) > 0 THEN ('Friday' + CASE WHEN (@freq_interval & (64)) > 0 THEN ', ' ELSE '' END) ELSE '' END + 
            CASE WHEN (@freq_interval & 64) > 0 THEN 'Saturday' ELSE '' END + ' '
        WHEN 16 THEN 'Occurs ' +
            CASE @freq_recurrence_factor
                WHEN 0 THEN 'monthly '
                WHEN 1 THEN 'every month '
                ELSE 'every ' + CONVERT(VARCHAR, @freq_recurrence_factor) + ' months '
            END +
            'on ' + 
            CASE
                WHEN @freq_interval = 0 THEN 'any day '
                ELSE 'day ' + CONVERT(VARCHAR, @freq_interval) + ' of the month '
            END
        WHEN 32 THEN 'Occurs ' +
            CASE @freq_recurrence_factor
                WHEN 0 THEN 'monthly '
                WHEN 1 THEN 'every month '
                ELSE 'every ' + CONVERT(VARCHAR, @freq_recurrence_factor) + ' months '
            END +
            'on ' + 
            CASE @freq_relative_interval
                WHEN 1 THEN 'First '
                WHEN 2 THEN 'Second '
                WHEN 4 THEN 'Third '
                WHEN 8 THEN 'Fourth '
                WHEN 16 THEN 'Last '
                ELSE ''
            END +
            CASE @freq_interval
                WHEN 1 THEN 'Sunday '
                WHEN 2 THEN 'Monday '
                WHEN 3 THEN 'Tuesday '
                WHEN 4 THEN 'Wednesday '
                WHEN 5 THEN 'Thursday '
                WHEN 6 THEN 'Friday '
                WHEN 7 THEN 'Saturday '
                WHEN 8 THEN 'Day '
                WHEN 9 THEN 'Weekday '
                WHEN 10 THEN 'Weekend Day '
                ELSE ''
            END
    END;

    -- Subday frequency logic
    IF @freq_subday_type IN (0, 1, 2, 4, 8)
        SET @schedule_description = @schedule_description +
            CASE @freq_subday_type
                WHEN 0 THEN 'at any time'
                WHEN 1 THEN 'at ' + FORMAT(CONVERT(DATETIME, @active_start_time_TIME), @time_format) -- active_start_time_TIME is not null always
                WHEN 2 THEN 'every ' + CONVERT(VARCHAR, @freq_subday_interval) + ' seconds'
                WHEN 4 THEN 'every ' + CONVERT(VARCHAR, @freq_subday_interval) + ' minutes'
                WHEN 8 THEN 'every ' + CONVERT(VARCHAR, @freq_subday_interval) + ' hours'
            END;
			
    IF @freq_subday_type <> 1
    BEGIN
        IF @active_start_time <> 0 AND @active_end_time <> 0
            SET @schedule_description = @schedule_description +
            ' between ' + FORMAT(CONVERT(DATETIME, @active_start_time_TIME), @time_format) + ' and ' + FORMAT(CONVERT(DATETIME, @active_end_time_TIME), @time_format);
        ELSE IF @active_start_time <> 0
            SET @schedule_description = @schedule_description +
            ' starting from ' + FORMAT(CONVERT(DATETIME, @active_start_time_TIME), @time_format);
        ELSE IF @active_end_time <> 0
            SET @schedule_description = @schedule_description +
            ' until ' + FORMAT(CONVERT(DATETIME, @active_end_time_TIME), @time_format);
    END;

    SET @schedule_description = @schedule_description + '.';

    IF @active_start_date_DATE IS NOT NULL OR @active_end_date_DATE IS NOT NULL
    BEGIN
        SET @schedule_description = @schedule_description + ' Schedule is used ';

        IF @active_start_date_DATE IS NULL AND @active_end_date_DATE IS NOT NULL
            SET @schedule_description = @schedule_description + 'until ' + FORMAT(CONVERT(DATETIME, @active_end_date_DATE), @date_format) + '.';
        ELSE IF @active_start_date_DATE IS NOT NULL AND @active_end_date_DATE IS NULL
            SET @schedule_description = @schedule_description + 'starting from ' + FORMAT(CONVERT(DATETIME, @active_start_date_DATE), @date_format) + '.';
        ELSE
            SET @schedule_description = @schedule_description + 'between ' + FORMAT(CONVERT(DATETIME, @active_start_date_DATE), @date_format) + ' and ' + FORMAT(CONVERT(DATETIME, @active_end_date_DATE), @date_format) + '.';
    END;

	RETURN;
END;