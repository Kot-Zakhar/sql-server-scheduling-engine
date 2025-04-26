CREATE PROCEDURE [dbo].[sp_get_schedule_next_execution_date_and_time] 
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
    @last_execution_date INT,
    @last_execution_time INT,
    @next_execution_date INT OUTPUT,
    @next_execution_time INT OUTPUT
)
AS
BEGIN
/*
If there is no next rundate/time, the function returns NULL.
Function works in local system time.

freq_type - int

	How frequently a job runs for this schedule.

	1 = One time only
	4 = Daily
	8 = Weekly
	16 = Monthly
	(not supported yet)32 = Monthly, relative to freq_interval
	(not supported)64 = Runs when the SQL Server Agent service starts
	(not supported)128 = Runs when the computer is idle

freq_interval - int	

	Days that the job is executed. Depends on the value of freq_type. 
	The default value is 0, which indicates that freq_interval is unused. 
	See the table below for the possible values and their effects.

freq_subday_type - int

	Units for the freq_subday_interval. The following are the possible 
	values and their descriptions.
	
	0 : Any time between the start time and end time
	1 : At the specified time (start time)
	2 : Seconds
	4 : Minutes
	8 : Hours

freq_subday_interval - int

	Number of freq_subday_type periods to occur between each execution of the job.

freq_relative_interval - int	

	When freq_interval occurs in each month, if freq_type is 32 (monthly relative). 
	Can be one of the following values:

	0 = freq_relative_interval is unused
	1 = First
	2 = Second
	4 = Third
	8 = Fourth
	16 = Last

freq_recurrence_factor - int

	Number of weeks or months between the scheduled execution of a job. 
	freq_recurrence_factor is used only if freq_type is 8, 16, or 32. 
	If this column contains 0, freq_recurrence_factor is unused.

active_start_date - int

	Date on which execution of a job can begin. The date is formatted as YYYYMMDD. NULL indicates today's date.

active_end_date - int

	Date on which execution of a job can stop. The date is formatted YYYYMMDD.

active_start_time - int

	Time on any day between active_start_date and active_end_date that job begins executing. 
	Time is formatted HHMMSS, using a 24-hour clock.

active_end_time - int

	Time on any day between active_start_date and active_end_date that job stops executing. 
	Time is formatted HHMMSS, using a 24-hour clock.

Value of freq_type				Effect on freq_interval
-------------------------------------------------------
	1 (once)					freq_interval is unused (0)
	4 (daily)					Every freq_interval days
	8 (weekly)					freq_interval is one or more of the following:
									1 = Sunday
									2 = Monday
									4 = Tuesday
									8 = Wednesday
									16 = Thursday
									32 = Friday
									64 = Saturday
	16 (monthly)				On the freq_interval day of the month
	32 (monthly, relative)		freq_interval is one of the following:
									1 = Sunday
									2 = Monday
									3 = Tuesday
									4 = Wednesday
									5 = Thursday
									6 = Friday
									7 = Saturday
									8 = Day
									9 = Weekday
									10 = Weekend day
	64 (starts when SQL Server Agent service starts)	freq_interval is unused (0)
	128 (runs when computer is idle)					freq_interval is unused (0)
*/

    SET NOCOUNT ON;

    -- Declare helper variables
	
	DECLARE @current_datetime DATETIME = CURRENT_TIMESTAMP;

    DECLARE @current_date INT = CONVERT(INT, FORMAT(@current_datetime, 'yyyyMMdd'));
    DECLARE @current_time INT = CONVERT(INT, FORMAT(@current_datetime, 'HHmmss'));
    DECLARE @tmp_current_date_DATE DATE;
    DECLARE @tmp_datefirst INT
    DECLARE @tmp_start_date_DATE DATE;

    -- Initialize output variables
    SET @next_execution_date = NULL;
    SET @next_execution_time = NULL;

	-- Nulls to defaults
    if (TRY_CONVERT(DATETIME, CONVERT(VARCHAR(8), @active_start_date)) IS NULL OR @active_start_date <= 0)
        SET @active_start_date = @current_date;
    if (TRY_CONVERT(TIME, STUFF(STUFF(RIGHT('000000' + CONVERT(VARCHAR(6), @active_start_time), 6), 5, 0, ':'), 3, 0, ':')) IS NULL OR @active_start_time <= 0)
        SET @active_start_time = 0;
    if (TRY_CONVERT(DATETIME, CONVERT(VARCHAR(8), @active_end_date)) IS NULL OR @active_end_date <= 0)
        SET @active_end_date = 99991231;
    if (TRY_CONVERT(TIME, STUFF(STUFF(RIGHT('000000' + CONVERT(VARCHAR(6), @active_end_time), 6), 5, 0, ':'), 3, 0, ':')) IS NULL OR @active_end_time <= 0)
        SET @active_end_time = 235959;
		
	-- Validation
	IF (@active_end_date < @active_start_date)
		OR (@active_end_time < @active_start_time)
		RETURN;

    IF [dbo].[agent_datetime](@last_execution_date, @last_execution_time) > [dbo].[agent_datetime](@current_date, @current_time)
    BEGIN
        SET @last_execution_date = NULL;
        SET @last_execution_time = NULL;
    END;

    IF @freq_type NOT IN (1, 4, 8, 16, 32)
        RETURN;

        
    -- Jumping into future to active_date/time
    -- if we are in active date range - adjust time
    IF @active_start_date <= @current_date AND @current_date <= @active_end_date
    BEGIN 
        -- if active time has passed for today - lets jumpt into tomorrow or active_start_date
        IF (@active_end_time < @current_time)
        BEGIN
            SET @current_date = @current_date + 1;
            SET @current_time = @active_start_time;
        END;

        -- if active time is in future - lets jump into future
        IF (@current_time < @active_start_time)
            SET @current_time = @active_start_time;
    END
    -- if date range is in future - jump into date range to the active_start_time
    ELSE IF @current_date < @active_start_date
    BEGIN
        SET @current_date = @active_start_date;
        SET @current_time = @active_start_time;
    END;
	
    SET @current_datetime = [dbo].[agent_datetime](@current_date, @current_time);

    -- if active perioud has been missed - no execution
    IF [dbo].[agent_datetime](@active_end_date, @active_end_time) < @current_datetime
        RETURN;
        
    -- Subday frequency logic
    IF @freq_subday_type > 0
    BEGIN
        IF @freq_subday_type = 1 -- At the specified time only (active_start_time)
        BEGIN
            IF @active_start_time < @current_time -- so if active_start_time was missed - schedule for tomorrow
            BEGIN
                SET @current_date = @current_date + 1;
                SET @current_time = @active_start_time;
                SET @current_datetime = [dbo].[agent_datetime](@current_date, @current_time);
            END;
        END
        ELSE IF @freq_subday_type in (2, 4, 8)
            AND @freq_subday_interval > 0
        BEGIN
            DECLARE @tmp_current_datetime DATETIME;
            IF @last_execution_date = @current_date AND @last_execution_time IS NOT NULL
			BEGIN            
                SET @tmp_current_datetime = [dbo].[agent_datetime](@last_execution_date, @last_execution_time);

                -- we need to add interval before the loop to prevent last_execution_time from matching current_time (later known as next_execution_time)
                SET @tmp_current_datetime = CASE
                    WHEN @freq_subday_type = 2 THEN DATEADD(SECOND, @freq_subday_interval, @tmp_current_datetime) -- seconds
                    WHEN @freq_subday_type = 4 THEN DATEADD(MINUTE, @freq_subday_interval, @tmp_current_datetime) -- minutes
                    WHEN @freq_subday_type = 8 THEN DATEADD(HOUR, @freq_subday_interval, @tmp_current_datetime) -- hours
                END;
			END
            ELSE
			BEGIN
                -- There is no sense to start from active_start_date+active_start_time
                -- because this happens within current day.
                -- Therefore, we can start from current_date+active_start_time
                SET @tmp_current_datetime = [dbo].[agent_datetime](@current_date, @active_start_time);
			END;

            WHILE @tmp_current_datetime < @current_datetime
            BEGIN
                SET @tmp_current_datetime = CASE
                    WHEN @freq_subday_type = 2 THEN DATEADD(SECOND, @freq_subday_interval, @tmp_current_datetime) -- seconds
                    WHEN @freq_subday_type = 4 THEN DATEADD(MINUTE, @freq_subday_interval, @tmp_current_datetime) -- minutes
                    WHEN @freq_subday_type = 8 THEN DATEADD(HOUR, @freq_subday_interval, @tmp_current_datetime) -- hours
                END;
            END;

            IF CONVERT(DATE, @tmp_current_datetime) > CONVERT(DATE, CONVERT(VARCHAR(8), @current_date))
				OR CONVERT(TIME, @tmp_current_datetime) > CONVERT(TIME, STUFF(STUFF(RIGHT('000000' + CONVERT(VARCHAR(6), @active_end_time), 6), 5, 0, ':'), 3, 0, ':'))
				-- current day / active period has finished, subday logic is no more applicable
            BEGIN
                SET @current_date = @current_date + 1;
                SET @current_time = @active_start_time;
            END
            ELSE
            BEGIN
                SET @current_time = CONVERT(INT, FORMAT(@tmp_current_datetime, 'HHmmss'));
            END;
            
            SET @current_datetime = [dbo].[agent_datetime](@current_date, @current_time);
        END;
    END;

    -- if active perioud has been missed - no execution
    IF [dbo].[agent_datetime](@active_end_date, @active_end_time) < @current_datetime
        RETURN;
		
    IF @freq_type = 1 -- One-time schedule
    BEGIN
		IF (@last_execution_date IS NOT NULL OR @last_execution_time IS NOT NULL)
			-- one-time schedule should not be executed anymore
			RETURN;
    
		SET @next_execution_date = @current_date;
		SET @next_execution_time = @current_time;

        RETURN;
    END;

    -- Function to calculate the next valid date based on frequency type
    IF @freq_type = 4 -- Daily
    BEGIN
        DECLARE @day_interval INT = @freq_interval;
        IF @freq_interval <= 0
            SET @day_interval = 1;
		-- start counting from either last_execution_date or active_start_date
        IF @last_execution_date IS NOT NULL
        BEGIN
            SET @tmp_current_date_DATE = CONVERT(DATE, CONVERT(varchar(8), @last_execution_date));

			-- If there is subday frequency, then last_execution_date may match next_execution_date
			-- otherwise we need to add interval before the loop to prevent last_execution_date from matching next_execution_date
			IF @freq_subday_type <= 0
				SET @tmp_current_date_DATE = DATEADD(DAY, @day_interval, @tmp_current_date_DATE);
        END
        ELSE
        BEGIN
            SET @tmp_current_date_DATE = TRY_CONVERT(DATE, CONVERT(varchar(8), @active_start_date));
        END;
        
        WHILE @tmp_current_date_DATE < TRY_CONVERT(DATE, CONVERT(varchar(8), @current_date))
        BEGIN
            SET @tmp_current_date_DATE = DATEADD(DAY, @day_interval, @tmp_current_date_DATE);
        END;

        SET @next_execution_date = TRY_CONVERT(INT, FORMAT(@tmp_current_date_DATE, 'yyyyMMdd'));

        -- active period is checked at the end of the function
    END
    ELSE IF @freq_type = 8 -- Weekly
    BEGIN
        SET @tmp_datefirst = @@DATEFIRST;
        SET DATEFIRST 7; -- setting Sunday as 1st day of week (default value)
		
        -- start counting from either last_execution_date or active_start_date
        IF @last_execution_date IS NOT NULL
        BEGIN
            SET @tmp_current_date_DATE = CONVERT(DATE, CONVERT(varchar(8), @last_execution_date));

            SET @tmp_start_date_DATE = @tmp_current_date_DATE;

			-- If there is subday frequency, then last_execution_date may match next_execution_date
			-- otherwise we need to add one day/freq_recurrence_factor before the loop to prevent last_execution_date from matching next_execution_date
			IF @freq_subday_type <= 0
            BEGIN
                SET @tmp_current_date_DATE = DATEADD(DAY, 1, @tmp_current_date_DATE);
            END;
        END
        ELSE
        BEGIN
            SET @tmp_current_date_DATE = TRY_CONVERT(DATE, CONVERT(varchar(8), @active_start_date));

            SET @tmp_start_date_DATE = @tmp_current_date_DATE;
        END;

        DECLARE @day_of_week INT = DATEPART(WEEKDAY, @tmp_current_date_DATE); -- 1 to 7 for Sunday to Saturday

        WHILE (@tmp_current_date_DATE < TRY_CONVERT(DATE, @current_datetime))
            OR (@freq_interval > 0 AND @freq_interval & POWER(2, @day_of_week - 1) = 0)
            OR (@freq_recurrence_factor > 0 AND (DATEPART(WEEK, @tmp_current_date_DATE) - DATEPART(WEEK, @tmp_start_date_DATE) < @freq_recurrence_factor))
        BEGIN
            SET @tmp_current_date_DATE = DATEADD(DAY, 1, @tmp_current_date_DATE);
            SET @day_of_week = CASE WHEN @day_of_week = 7 THEN 1 ELSE @day_of_week + 1 END;
        END;
        
        SET @next_execution_date = TRY_CONVERT(INT, FORMAT(@tmp_current_date_DATE, 'yyyyMMdd'));

        SET DATEFIRST @tmp_datefirst;
        -- active period is checked at the end of the function
    END
    ELSE IF @freq_type = 16 -- Monthly
    BEGIN
        SET @tmp_datefirst = @@DATEFIRST;
        SET DATEFIRST 7; -- setting Sunday as 1st day of week (default value)

        -- start counting from either last_execution_date or active_start_date
        IF @last_execution_date IS NOT NULL
        BEGIN
            SET @tmp_current_date_DATE = CONVERT(DATE, CONVERT(varchar(8), @last_execution_date));

			-- If there is subday frequency, then last_execution_date may match next_execution_date
			-- otherwise we need to add days before the loop to prevent last_execution_date from matching next_execution_date
			IF @freq_subday_type <= 0
            BEGIN
                SET @tmp_start_date_DATE = @tmp_current_date_DATE;
                -- Jump to the 1st day of next month to last_execution_date's month
                SET @tmp_current_date_DATE = DATEFROMPARTS(DATEPART(YEAR, @tmp_current_date_DATE), DATEPART(MONTH, @tmp_current_date_DATE), 1);
                SET @tmp_current_date_DATE = DATEADD(MONTH, 1, @tmp_current_date_DATE);
            END;
        END
        ELSE
        BEGIN
            SET @tmp_current_date_DATE = TRY_CONVERT(DATE, CONVERT(varchar(8), @active_start_date));
        END;
        
        WHILE (@tmp_current_date_DATE < TRY_CONVERT(DATE, @current_datetime))
            OR (@freq_interval > 0 AND DAY(@tmp_current_date_DATE) <> @freq_interval)
            OR (@freq_recurrence_factor > 0 AND @tmp_start_date_DATE IS NOT NULL AND (DATEPART(MONTH, @tmp_current_date_DATE) - DATEPART(MONTH, @tmp_start_date_DATE) < @freq_recurrence_factor))
        BEGIN
            SET @tmp_current_date_DATE = DATEADD(DAY, 1, @tmp_current_date_DATE);
        END;

        SET @next_execution_date = TRY_CONVERT(INT, FORMAT(@tmp_current_date_DATE, 'yyyyMMdd'));

        SET DATEFIRST @tmp_datefirst;
        -- active period is checked at the end of the function
    END
    ELSE IF @freq_type = 32 -- Monthly Relative
    BEGIN
        SET @tmp_datefirst = @@DATEFIRST;
        SET DATEFIRST 7; -- setting Sunday as 1st day of week (default value)

        -- start counting from either last_execution_date or active_start_date
        IF @last_execution_date IS NOT NULL
        BEGIN
            SET @tmp_current_date_DATE = CONVERT(DATE, CONVERT(varchar(8), @last_execution_date));
            
            -- If there is subday frequency, then last_execution_date may match next_execution_date
            -- otherwise we need to add days before the loop to prevent last_execution_date from matching next_execution_date
            IF @freq_subday_type <= 0
            BEGIN
                SET @tmp_start_date_DATE = @tmp_current_date_DATE;
                -- Jump to the 1st day of next month to last_execution_date's month
                SET @tmp_current_date_DATE = DATEFROMPARTS(DATEPART(YEAR, @tmp_current_date_DATE), DATEPART(MONTH, @tmp_current_date_DATE), 1);
                SET @tmp_current_date_DATE = DATEADD(MONTH, 1, @tmp_current_date_DATE);
            END;
        END
        ELSE
        BEGIN
            SET @tmp_current_date_DATE = TRY_CONVERT(DATE, CONVERT(varchar(8), @active_start_date));
			SET @tmp_current_date_DATE = DATEFROMPARTS(DATEPART(YEAR, @tmp_current_date_DATE), DATEPART(MONTH, @tmp_current_date_DATE), 1);
        END;
        
        -- If we're starting from last execution, calculate the target month based on recurrence factor
        IF @tmp_start_date_DATE IS NOT NULL
        BEGIN
            -- Calculate months between start date and current date
            DECLARE @months_between INT = DATEDIFF(MONTH, @tmp_start_date_DATE, @tmp_current_date_DATE);
            
            -- Calculate how many recurrence periods we need to move forward
            DECLARE @recurrence_periods INT = CEILING((@months_between + 1.0) / @freq_recurrence_factor);
            
            -- Calculate the target date by moving forward the required number of months
            SET @tmp_current_date_DATE = DATEADD(MONTH, @recurrence_periods * @freq_recurrence_factor, @tmp_start_date_DATE);
            SET @tmp_current_date_DATE = DATEFROMPARTS(YEAR(@tmp_current_date_DATE), MONTH(@tmp_current_date_DATE), 1);
        END;
        
        -- For Monthly Relative, @freq_relative_interval represents which occurrence (1=first, 2=second, 4=third, 8=fourth, 16=last)
        -- @freq_interval represents which day (1=Sunday, 2=Monday, ..., 7=Saturday, 8=Day, 9=Weekday, 10=Weekend day)
        -- @freq_recurrence_factor represents how many months to skip
        
        DECLARE @target_occurrence INT = @freq_relative_interval;
        DECLARE @target_day_type INT = @freq_interval;
        DECLARE @months_to_skip INT = CASE WHEN @freq_recurrence_factor IS NULL OR @freq_recurrence_factor = 0 THEN 1 ELSE @freq_recurrence_factor END;
        DECLARE @current_month INT;
        DECLARE @target_month INT;
        DECLARE @target_date DATE;
        DECLARE @weekend_count INT = 1;
        DECLARE @weekday_count INT = 1;
        DECLARE @first_run BIT = 1;
        
        -- Find the target date based on occurrence and day type
        SET @target_date = @tmp_current_date_DATE;
        
        WHILE @first_run = 1 OR @target_date < TRY_CONVERT(DATE, @current_datetime)
        BEGIN
            IF @first_run = 0
            BEGIN
		        -- If the target date is in the past, move to the next occurrence
                -- Move to the first day of the next month based on recurrence factor
                SET @tmp_current_date_DATE = DATEADD(MONTH, @months_to_skip, @tmp_current_date_DATE);
                SET @tmp_current_date_DATE = DATEFROMPARTS(YEAR(@tmp_current_date_DATE), MONTH(@tmp_current_date_DATE), 1);

                -- Find the target date based on occurrence and day type
                SET @target_date = @tmp_current_date_DATE;
            END
            ELSE
            BEGIN
                SET @first_run = 0;
            END;

            -- Handle different day types
            IF @target_day_type BETWEEN 1 AND 7 -- Specific day of week
            BEGIN
                -- Find the first occurrence of the specified day in the month
                WHILE DATEPART(WEEKDAY, @target_date) <> @target_day_type
                BEGIN
                    SET @target_date = DATEADD(DAY, 1, @target_date);
                END;
                
                -- Move to the specified occurrence

                -- Skip @target_occurrence = 1, because it's already at the first occurrence
                
                IF @target_occurrence IN (2, 4, 8) -- Second, third or fourth occurrence
                BEGIN
                    SET @target_date = DATEADD(WEEK, 
                        CASE @target_occurrence
                            WHEN 2 THEN 1
                            WHEN 4 THEN 2
                            WHEN 8 THEN 3
                        END,
                        @target_date
                    );
                    
                    -- If we've moved to the next month, go back to the last occurrence in the previous month
                    IF DATEPART(MONTH, @target_date) <> DATEPART(MONTH, @tmp_current_date_DATE)
                    BEGIN
                        -- Go back to the last day of the previous month
                        SET @target_date = EOMONTH(@tmp_current_date_DATE);
                        
                        -- Find the last occurrence of the specified day in the month
                        WHILE DATEPART(WEEKDAY, @target_date) <> @target_day_type
                        BEGIN
                            SET @target_date = DATEADD(DAY, -1, @target_date);
                        END;
                    END;
                END
                ELSE IF @target_occurrence = 16 -- Last occurrence
                BEGIN
                    -- Start from the end of the month and go backwards
                    SET @target_date = EOMONTH(@tmp_current_date_DATE);
                    
                    -- Find the last occurrence of the specified day in the month
                    WHILE DATEPART(WEEKDAY, @target_date) <> @target_day_type
                    BEGIN
                        SET @target_date = DATEADD(DAY, -1, @target_date);
                    END;
                END;
            END
            ELSE IF @target_day_type = 8 -- Day
            BEGIN
                -- Skip @target_occurrence = 1, because it's already at the first occurrence
                IF @target_occurrence = 2 -- Second day
                BEGIN
                    SET @target_date = DATEADD(DAY, 1, @tmp_current_date_DATE);
                END
                ELSE IF @target_occurrence = 4 -- Third day
                BEGIN
                    SET @target_date = DATEADD(DAY, 2, @tmp_current_date_DATE);
                END
                ELSE IF @target_occurrence = 8 -- Fourth day
                BEGIN
                    SET @target_date = DATEADD(DAY, 3, @tmp_current_date_DATE);
                END
                ELSE IF @target_occurrence = 16 -- Last day
                BEGIN
                    SET @target_date = EOMONTH(@tmp_current_date_DATE);
                END;
            END
            ELSE IF @target_day_type = 9 -- Weekday
            BEGIN
                -- Find the first weekday in the month
                WHILE DATEPART(WEEKDAY, @target_date) = 1 OR DATEPART(WEEKDAY, @target_date) = 7 -- Sunday or Saturday
                BEGIN
                    SET @target_date = DATEADD(DAY, 1, @target_date);
                END;
                
                -- Skip @target_occurrence = 1, because it's already at the first occurrence

                IF @target_occurrence IN (2, 4, 8) -- Second, third or fourth weekday
                BEGIN
                    SET @weekday_count = 1;
                    WHILE @weekday_count < CASE @target_occurrence
                        WHEN 2 THEN 2
                        WHEN 4 THEN 3
                        WHEN 8 THEN 4
                    END
                    BEGIN
                        SET @target_date = DATEADD(DAY, 1, @target_date);
                        -- Skip weekends
                        WHILE DATEPART(WEEKDAY, @target_date) = 1 OR DATEPART(WEEKDAY, @target_date) = 7
                        BEGIN
                            SET @target_date = DATEADD(DAY, 1, @target_date);
                        END;
                        SET @weekday_count = @weekday_count + 1;
                    END;
                END
                ELSE IF @target_occurrence = 16 -- Last weekday
                BEGIN
                    SET @target_date = EOMONTH(@tmp_current_date_DATE);
                    WHILE DATEPART(WEEKDAY, @target_date) = 1 OR DATEPART(WEEKDAY, @target_date) = 7
                    BEGIN
                        SET @target_date = DATEADD(DAY, -1, @target_date);
                    END;
                END;
                
                -- If we've moved to the next month, go back to the last weekday in the previous month
                IF DATEPART(MONTH, @target_date) <> DATEPART(MONTH, @tmp_current_date_DATE)
                BEGIN
                    SET @target_date = EOMONTH(@tmp_current_date_DATE);
                    WHILE DATEPART(WEEKDAY, @target_date) = 1 OR DATEPART(WEEKDAY, @target_date) = 7
                    BEGIN
                        SET @target_date = DATEADD(DAY, -1, @target_date);
                    END;
                END;
            END
            ELSE IF @target_day_type = 10 -- Weekend day
            BEGIN
                -- Find the first weekend day in the month
                WHILE DATEPART(WEEKDAY, @target_date) <> 1 AND DATEPART(WEEKDAY, @target_date) <> 7 -- Not Sunday and not Saturday
                BEGIN
                    SET @target_date = DATEADD(DAY, 1, @target_date);
                END;
                
                -- Skip @target_occurrence = 1, because it's already at the first occurrence

                IF @target_occurrence IN (2, 4, 8) -- Second, third or fourth weekend day
                BEGIN
                    SET @weekend_count = 1;
                    WHILE @weekend_count < CASE @target_occurrence 
                        WHEN 2 THEN 2
                        WHEN 4 THEN 3
                        WHEN 8 THEN 4
                    END
                    BEGIN
                        SET @target_date = DATEADD(DAY, 1, @target_date);
                        -- Skip weekdays
                        WHILE DATEPART(WEEKDAY, @target_date) <> 1 AND DATEPART(WEEKDAY, @target_date) <> 7
                        BEGIN
                            SET @target_date = DATEADD(DAY, 1, @target_date);
                        END;
                        SET @weekend_count = @weekend_count + 1;
                    END;
                END
                ELSE IF @target_occurrence = 16 -- Last weekend day
                BEGIN
                    SET @target_date = EOMONTH(@tmp_current_date_DATE);
                    WHILE DATEPART(WEEKDAY, @target_date) <> 1 AND DATEPART(WEEKDAY, @target_date) <> 7
                    BEGIN
                        SET @target_date = DATEADD(DAY, -1, @target_date);
                    END;
                END;
                
                -- If we've moved to the next month, go back to the last weekend day in the previous month
                IF DATEPART(MONTH, @target_date) <> DATEPART(MONTH, @tmp_current_date_DATE)
                BEGIN
                    SET @target_date = EOMONTH(@tmp_current_date_DATE);
                    WHILE DATEPART(WEEKDAY, @target_date) <> 1 AND DATEPART(WEEKDAY, @target_date) <> 7
                    BEGIN
                        SET @target_date = DATEADD(DAY, -1, @target_date);
                    END;
                END;
            END;
        END;
        
        SET @tmp_current_date_DATE = @target_date;
        SET @next_execution_date = TRY_CONVERT(INT, FORMAT(@tmp_current_date_DATE, 'yyyyMMdd'));

        SET DATEFIRST @tmp_datefirst;
        -- active period is checked at the end of the function
    END;

	-- freq_type = 64 NOT SUPPORTED
	-- Automatically starts when SQLServerAgent starts

    -- freq_type = 128 NOT SUPPORTED
    -- runs when computer is idle

    -- if next date is today
    -- schedule at current time to utilize subday functionality
    -- otherwise at active_start_time to override subday calculations
    IF @next_execution_date = @current_date
        SET @next_execution_time = @current_time;
    ELSE
        SET @next_execution_time = @active_start_time;
        
    -- Check if the calculated next execution date and time is within the active period
    IF @active_end_date < @next_execution_date
        OR (@active_end_date = @next_execution_date AND @active_end_time < @next_execution_time)
    BEGIN
        SET @next_execution_date = NULL;
        SET @next_execution_time = NULL;
    END;
END;
