# VBA-Challenge
VBA script to analyze stock.

For this challenge I had difficulty setting up the logic. I worked with BCS Learning Assistant on 3/10/24 to help set up the logic for my three If statements in my For loop. From this session I was able to develop the skill of designing my code before I start.

I also had difficulty debugging my code on 3/9 and 3/10 and worked with BCS Learning Assistants. From these sessions I learned that the order of defining variables can cause errors. For examples my LastRowK was assigned up at the very top but at that point in time column K had no data therefore I needed to assign the variable after the column had data.

Instructions for future reference:
Create a script that loops through all the stocks for one year and outputs the following information:

The ticker symbol

Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

The total stock volume of the stock. 

Add functionality to your script to return the stock with the "Greatest % increase", 
"Greatest % decrease", and "Greatest total volume".

Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

Make sure to use conditional formatting that will highlight positive change in green and negative change in red.