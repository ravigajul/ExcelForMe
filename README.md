# ExcelForMe
    Excel calculates Dates based on 1/1/1900 --42379 is the 42379 days from 1/1/1900.
    09:00 + 10:00 will be some number and not an exact sum of the time due to above logic. Format it to [hh]:mm to show exact sum.
# Text to columns by delimiter.
# Index(Small(if(),row()))
    Ex: Array formulas (CSE -->Ctrl +Shift+Enter)
    {=IFERROR(INDEX($A$1:$F$14,SMALL(IF($A$1:$F$14=$I$2,ROW($A$1:$F$14)),ROW(1:1)),COLUMN(A:A)),"")}
# Vlookup(match()<<instead of column numbers >>) for dynamic column search
    =VLOOKUP($I$2,$A$1:$F$14,MATCH("Full Name",$A$1:$F$1,0))
# Choose (for looking backwards)
    =VLOOKUP(E3,CHOOSE({1,2,3},C3:C7,B3:B7,A3:A7),3,FALSE)
# Index(match) for looking any direction
# CountA --returns the count of non empty rows
# Dynamic name range for data
    =OFFSET($A$1,0,0,COUNTA($A:$A),COUNTA($1:$1))
    Or 
    ==OFFSET(Sheet1!$A$1,MATCH(Sheet2!$A$1,Sheet1!$A:$A,0)-1,2,1,3)
# Choose---like switch case
	=CHOOSE(MONTH(A4),L4,L5,O4,L6,,L7,L8,L10,L9,L11,L12,L13,L14,L15)
# Find
# RandBetween
# Proper --to capitalize first letter of every word in a sentence
# LEFT,RIGHT,MID,LEN
# indirect
# To return multiple matches 
    https://www.youtube.com/watch?v=fDB1Ktyhp3
    =AGGREGATE(15,3,(SurveyData!$C1:$C4=Sheet1!$C$2)/(SurveyData!$C1:$C4=Sheet1!$C$2)*(ROW('SurveyData'!C1:C4)-ROW(SurveyData!C1)),ROWS(Sheet1!$D$3:D4))
    =INDEX(SurveyData!$B:$P,AGGREGATE(15,3,(SurveyData!$C1:$C4=Sheet1!$C$2)/(SurveyData!$C1:$C4=Sheet1!$C$2)*(ROW('SurveyData'!C1:C4)-ROW(SurveyData!C1)),ROWS($E$4:F4)),1)
    Note : Aggregate is to avoid CSE ( Control + Shift+ Enter). ROWS is to get the rows(1,2,3) dynamically
    Final formula
    =IF(ROWS($E$4:F4)<=$B$2,INDEX(SurveyData!$B:$P,AGGREGATE(15,3,(SurveyData!$C:$C=Helper!$C$2)/(SurveyData!$C:$C=Helper!$C$2)*(ROW(SurveyData!H:H)),ROWS($E$4:F4)),2),"")
    =Sequence(53,7,1,1) Generates a sequence  of numbers for 53 rows and 7 columns
