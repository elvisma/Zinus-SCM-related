# -------------------------------------------------------------------------------
# Name:        tip
# Purpose:     tip calculator
#
# Author:      Rula Khayrallah
#
# Created:     09/14/2014
# Copyright:   (c) Rula Khayrallah 2014
# -------------------------------------------------------------------------------

"""
Tip calculator assuming a 20% tip rate

Prompt the user for the cost of their meal.
Print the tip amount and the total cost.
"""
TIP_RATE=20/100
user_input=input("Please enter the cost of your meal in $:")
cost=float(user_input)
tip=TIP_RATE*cost
tip=round(tip,2)
print("Tip Amount: $", tip, sep="")
total=cost+tip
total=round(total,2)
print("Total Cost: $", total, sep='')


print("CS"+"A"+"21")



