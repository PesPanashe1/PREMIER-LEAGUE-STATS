**FINAL PL STANDINGS PROJECT NOTES**

**USING EXCEL TOOLS TO CREATE ALL PREMIER LEAGUE TABLES ACROSS THE 1992/93 TO 2023/24 SEASONS**.

1) Created 31 unformatted Worksheets.


2) Created a Macro and wrote VBA Procedure Code that adds and Boldens Headers, as well as 
   adding Position Numbers to every sheet, at the same time.
   It also includes Font, Cell Color and Adjustments to Cell Widths and Heights.


3) Used Flash Fill to enter number of Matches Played.


4) - Formula for finding Goal Difference number =(G2-H2)

   - Formula for finding Points =( (D2*3)+(E2*1)+(F2*0) ) 

   - Formula for adding the "+" and "-" Prefixes to the Goal Diffence Number =CONCATENATE(IF((G2-H2) > 0,"+","-"), G2-H2)
 
   - Macro for adding all 3 formulas to every sheet at the same time.



7) - Formula for finding Losses =( C2-(D2+E2) )
    
   - Macro for adding Losses Formula to every sheet at the same time. 


8) - Macro for adding number of matches played(38) to every sheet with 20 teams.
 

9) - Manually entered all Team Names based on their positions.
   - Created a Drop Down function that includes a list of all Teams that have participated in the Premier League since its inception.
     The purpose being to avoid the strenuous task of typing the same Team Names over and over again.     


10) Added Cell color to discern the Champion(Green) and Relegated teams(Red).


11) Macro that adds a 'CHAMPION' and 'RELEGATED' status to the respective teams across all worksheets.


12) Macro that Protects every worksheet from any editing by third-parties.
