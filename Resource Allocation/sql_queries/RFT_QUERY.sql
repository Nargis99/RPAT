-- STEP 1. Change the Database
USE INDIA_GTSG;

-- STEP 2. Drop existing table 
DROP TABLE Sheet1$;
DROP TABLE RFT_P6;

-- STEP 3. Import the RFT excel file 
SELECT TOP 10 *  FROM Sheet1$;                                                                           --to check the imported dataset 

-- STEP 4. Dump 'Sheet1$' data into RFT_P6 
SELECT * INTO RFT_P6 FROM Sheet1$;

-- STEP 5. To check if the date column header are in correct format
SELECT * FROM RFT_P6;

-- STEP 6. Deleting rows where "Resource (Utilization Target %)/Job Code" is 'Inactive
DELETE FROM RFT_P6 WHERE "Resource (Utilization Target %)/Job Code" IN
   (SELECT "Resource (Utilization Target %)/Job Code" FROM RFT_P6 WHERE "Resource (Utilization Target %)/Job Code" LIKE '%Inactive%');

-- STEP 7. Adding new column 'Project Number'
ALTER TABLE RFT_P6 ADD "Project Number" VARCHAR(255);

-- STEP 8. Updating column 'Project Number' 
UPDATE RFT_P6 set "Project Number"=CASE WHEN "Project" LIKE '%-%' THEN  RIGHT("Project", 
		 CHARINDEX(' ', (REVERSE("Project")))-1) ELSE 'NA' END

-- STEP 9. Add multiple columns in the table 
ALTER TABLE RFT_P6 
ADD Resource_Name VARCHAR(255),
    Employee_code VARCHAR(255),
	Utilization_Target INT,
	Utilization_Target_percent VARCHAR(255);

-- STEP 10. Updating Resource_Name
UPDATE RFT_P6 SET Resource_Name=CASE
       WHEN "Resource (Utilization Target %)/Job Code" LIKE '%-%' 
	   THEN LEFT("Resource (Utilization Target %)/Job Code", 
	   CHARINDEX(RIGHT("Resource (Utilization Target %)/Job Code", 
	   CHARINDEX('-', (REVERSE("Resource (Utilization Target %)/Job Code")))+1)
		 , (("Resource (Utilization Target %)/Job Code")))) ELSE "Resource (Utilization Target %)/Job Code"
       END 

-- STEP 11. Updating Utilization_Target_percent
UPDATE RFT_P6 SET Utilization_Target_percent=CASE WHEN "Resource (Utilization Target %)/Job Code" LIKE '%-%' THEN RIGHT("Resource (Utilization Target %)/Job Code", 
		 CHARINDEX('(', (REVERSE("Resource (Utilization Target %)/Job Code")))) ELSE 'NA' END

-- STEP 12. Updating column 'Utilization_Target'
UPDATE RFT_P6 SET Utilization_Target=CASE
         WHEN "Resource (Utilization Target %)/Job Code" LIKE '%-%' 
		 THEN dbo.UDF_ExtractNumbers(Utilization_Target_percent) 
ELSE 0 END

-- STEP 13. Updating column 'employee_code'
UPDATE RFT_P6 SET employee_code=CASE
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

-- STEP 14. Extract 'employee code'
UPDATE RFT_P6 SET employee_code=CASE
         WHEN "Resource (Utilization Target %)/Job Code" LIKE '%-%' 
		 THEN dbo.UDF_ExtractNumbers(employee_code) 
ELSE 0 END

---------------------------------------- QUERY ENDS HERE -----------------------------------------------


















-- Function to extract numbers from string of characters
/*Create function UDF_ExtractNumbers
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
End*/
