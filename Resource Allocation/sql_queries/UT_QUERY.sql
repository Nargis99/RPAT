-- STEP 1. To change the database
USE INDIA_GTSG;

-- STEP 2. To view the table
SELECT * FROM ut;

-- STEP 3. Comparing table UT and RFT_P6 to find resources which are not in UT and inserting in UT
INSERT INTO UT ("Resource (Utilization Target %)/Job Code")
SELECT DISTINCT T1."Resource (Utilization Target %)/Job Code" 
FROM RFT_P6 T1
LEFT JOIN UT T2
ON T1."Resource (Utilization Target %)/Job Code" = T2."Resource (Utilization Target %)/Job Code" WHERE T2."Resource (Utilization Target %)/Job Code" IS NULL;

-- to check distincts
--SELECT DISTINCT COUNT("Resource (Utilization Target %)/Job Code") from ut
-- to view table
--select count(*) from ut

-- STEP 4. Updating Resource_Name
UPDATE UT SET Resource_Name=CASE
         WHEN "Resource (Utilization Target %)/Job Code" LIKE '%-%' 
		 THEN left("Resource (Utilization Target %)/Job Code", 
		 CHARINDEX(right("Resource (Utilization Target %)/Job Code", 
		 CHARINDEX('-', (REVERSE("Resource (Utilization Target %)/Job Code")))+1)
		 , (("Resource (Utilization Target %)/Job Code")))) ELSE "Resource (Utilization Target %)/Job Code"
       END 

-- STEP 5. Updating Utilization_Target_percent
UPDATE UT SET Utilization_Target_percent=CASE WHEN "Resource (Utilization Target %)/Job Code" LIKE '%-%' THEN  right("Resource (Utilization Target %)/Job Code", 
		 CHARINDEX('(', (REVERSE("Resource (Utilization Target %)/Job Code")))) ELSE 'NA' END

-- STEP 6. Updating column Utilization_Target
UPDATE ut SET Utilization_Target=CASE
         WHEN "Resource (Utilization Target %)/Job Code" LIKE '%-%' 
		 THEN dbo.UDF_ExtractNumbers(Utilization_Target_percent) 
ELSE 0 END

-- STEP 7. To Extarct employee code
UPDATE ut SET employee_code=CASE
         WHEN "Resource (Utilization Target %)/Job Code" LIKE '%-%' 
		 THEN   
SUBSTRING("Resource (Utilization Target %)/Job Code",
LEN(LEFT("Resource (Utilization Target %)/Job Code",
CHARINDEX('-', "Resource (Utilization Target %)/Job Code")+1)),
LEN("Resource (Utilization Target %)/Job Code") - LEN(LEFT("Resource (Utilization Target %)/Job Code",
CHARINDEX('-', "Resource (Utilization Target %)/Job Code"))) - LEN(RIGHT("Resource (Utilization Target %)/Job Code",
CHARINDEX(' ', (REVERSE("Resource (Utilization Target %)/Job Code"))))))
         ELSE "Resource (Utilization Target %)/Job Code"
       END

-- STEP 8. To remove extra space
UPDATE ut SET employee_code=CASE
         WHEN "Resource (Utilization Target %)/Job Code" LIKE '%-%' 
		 THEN dbo.UDF_ExtractNumbers(employee_code) 
else 0 end

-- STEP 9. Delete inactive resource
delete from ut where "Utilization_Target_percent" LIKE '%Inactive%'

-- STEP 10. To remove duplicates
WITH table_nameCTE AS  
(  
   SELECT *, ROW_NUMBER() over (PARTITION BY "Resource_Name" ORDER BY "Resource_Name") as rk  
   FROM ut
)  
DELETE FROM table_nameCTE WHERE rk >1

---------------------------------------QUERY ENDS HERE-----------------------------------------














--function to extract numbers from string of characters
Create function UDF_ExtractNumbers
(  
  @input varchar(255)  
)  
Returns varchar(255)  
As  
Begin  
  Declare @alphabetIndex int = Patindex('%[^0-9]%', @input)  
  Begin  
    While @alphabetIndex > 0  
    Begin  
      Set @input = Stuff(@input, @alphabetIndex, 1, '' )  
      Set @alphabetIndex = Patindex('%[^0-9]%', @input )  
    End  
  End  
  Return @input
End
