# vba_challenge
enable developer mode
open editor
create code to run macros on all sheets in workbook. I listened to Instructor Manish explaining this part to a student during office hours.
for each sheet, first create the column headings.
define all the variable that will be used. the first 7 variable i got from stackoverflow. 
determine last row. i got this code from stackoverflow
start the loop from the 2nd row till the lastrow
create an if statment to have the loop go through the rows and pick each ticker that is different than the previous one and do the folowing: add the previous ticker's close price and calculate  its percent change value. i got these 2 codes from stackoverflow and edited them with the help from a programmer friend named Abbas. 
then using 3 nested if statments: find the greatest % increase, greatest % decrease, and the greatest total vol. i've created these 3 codes with the help from a programmer friend named Abbas. 
continue the code to print the previous ticker data to the summary table in columns I, J, K,and L as they were named previously. 
imbed a nested loop to conditionally format the cells in column J.
iterate th summary row with the new ticker data
then if the condition of the main if statment is false, just add up the previous tickers volum to column L
repeat the process by going to the next i
print the greatest % increase, greates % decrease, and greatest total volumn that we calculated in the previous loop in columns P and Q. 
format the values in the Q column by adding the % sign
