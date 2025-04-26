CREATE TABLE [dbo].[schedules] (
    schedule_id INT PRIMARY KEY IDENTITY,
    [name] NVARCHAR(255) NOT NULL,
    [description] NVARCHAR(255) NULL,
    freq_type INT NOT NULL,
    freq_interval INT NOT NULL,
    freq_subday_type INT NOT NULL,
    freq_subday_interval INT NOT NULL,
    freq_relative_interval INT NOT NULL,
    freq_recurrence_factor INT NOT NULL,
    active_start_date INT NOT NULL,
    active_end_date INT NOT NULL,
    active_start_time INT NOT NULL,
    active_end_time INT NOT NULL
);

CREATE TABLE [dbo].[next_checks] (
    schedule_id INT NOT NULL,
    FOREIGN KEY (schedule_id) REFERENCES [dbo].[schedules](schedule_id) ON DELETE CASCADE,
    previous_execution_date INT NULL,
    previous_execution_time INT NULL,
    next_execution_date INT NULL,
    next_execution_time INT NULL,
);