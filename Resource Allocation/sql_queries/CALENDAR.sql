-- STEP 1. To change Database
USE india_gtsg;

-- STEP 2. Drop the exisiting table 'calendar'
--DROP calendar;

-- STEP 3. Import calendar file

-- STEP 4. To view table
SELECT * FROM calendar;

-- STEP 5. To create new column
ALTER TABLE calendar ALTER COLUMN "Week No." INT;

sp_help calendar