# Stock-Analysis (2017 and 2018)
## Overview
- I am helping Steve analyze stocks for his parents and decide what stocks are worth investing in by comparing how the stocks did in 2017 and 2018. In order to achieve this, I found a way to formulate a button through macros that will give us the total daily volume and yearly return. Since a button was created, I edited the macro to run through all the stocks instead of a select few (12) in order to increase the user's avibility to analyze and greater amount of stocks at once.
- The purpose of this challenge is to recode the original VBA code in order to make it perform faster by optimizing the code while having it run efficiently.

## Results 
- The following are the timer results from the original Macros that were created through the module 2. These are 2017 and 2018 respectively.
 
![Original_VBACode_2017](./Resources/Original_VBACode_Time_2017.png)               ![Orignial_VBACode_2018](./Resources/Original_VBACode_Time_2018.png)

- The following pictures are from the new code that I refactored in order to make the run time faster, these pictures include the analysis for their respective years as well as the run time from the refactored VBA code. 

#### Refactored 2017 timer and 2017 Analysis
![Refactored_VBA_Code_2017](./Resources/VBA_Challenge_2017.png)

#### Refactored 2018 timer and 2018 Analysis
![Refactored_VBA_Code_2018](./Resources/VBA_Challenge_2018.png)

- Its great that my VBA code was able to run A LOT faster and smoother after it was refactored. However, the objective is to help my friend Steve figure out what stocks his parents should invest in. Now am I a little bias because I have a finance degree? Sure I am! Although through 2017 and 2018 only two stocks showed that they are profitable there is just not enough to state which stocks his parents should invest in. There are for sure a few stocks that already seem like a bad idea to invest in such as JKS and DQ but as for which stocks are worth it all depends on risk aversion and how much his parents are willing to lose in order to have a higher posible return. ENPH and RUN seem like good stocks for longterm hold, a few of these stocks show a high volatility which can scare off Steve's parents from investing in them.
