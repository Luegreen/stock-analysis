 #stock-analysis
 
 ### Overview of Project
 Using VBA, the objective was to create code that could analyse multiple stocks with the click of a button, aiding a hypothetical client in chosing a successful stock purchase for his parents. He was headed to DQ and, with the aid of this analysis, would be able to confirm that choice of stock option or make a better one. 
 
 ### Results
 **Comparing Performance between 2017 and 2018 for VBA Stocks:**
 
 The price of the 12 stocks reviewed experienced an overall increased over the year of 2017 and then decreased in 2018.
 Only 58% of the 12 companies reviewed experienced an increase in volume.
 
 Regarding the best recommendation for stock purchase, there is one clear winner but I will list the top 3 choices: 
 - ENPH experienced an 129.5% increase in 2017 and another 81.9% increase in 2018 AS WELL AS an almost 300% increase in volume. This makes it clearly the best stock in which to invest.
 - SEDG experienced an 184% increase in 2017 and then dropped only 7.8% in 2018. The volume sold increased around 15%. This puts it in second place as a stock worth investment.
 - RUN experienced a modest 5.5% increase in 2017 but picked up in 2018 with a 84% increase. This perhaps drove a vlume increase of nearly 100% which bodes well for further trading aspirations. 
 
 The chart below shows each of the 12 stocks, their volume for 2017 and 2018 including an indication of change (green circle = volume increase, red circle = volume decrease from 2017 to 2018, as well as % price increase/decrease (green = increase, red = decrease).
 
 
 <img width="622" alt="Screen Shot 2021-04-25 at 8 25 56 PM" src="https://user-images.githubusercontent.com/14239715/116016276-24126080-a60a-11eb-9020-bab657a15e12.png">
 
 
 
 **Execution time of original script and refactored script**
 
 On first analysis, the refactored script ran slower than my original script by .05 seconds for 2017 and .03 seconds for 2018. This was discouraging until I realized that the refactored script was also formatting the final graph. My original script had no formatting. I built the formatting script separately and added it into the GUI as a second button option (run the script by a button to analyse the stocks, then push another button to format your graph). So, to get an apples to apples comparison, I realized that I had to equalize the scripts. I then took the formatting out of the refactored script to time it's run, and also I added formatting to the orignial script, and timed it's run. The results are as follows:
 
 
 <img width="330" alt="Screen Shot 2021-04-25 at 9 16 20 PM" src="https://user-images.githubusercontent.com/14239715/116016957-41482e80-a60c-11eb-9c45-956a42b4c7cd.png">
 
 In this chart you can see that while refactoring did save time for the pre-formatted script (.05 seconds for 2017 and .07. seconds for 2018), the formatted scripts, both before and after refactoring had the same run time, therefor refactoring saved no time for formatted scripts. 



 
 ### Summary
 
  **Advantages and Disadvantages of refactoring code in general**
  
  Refactoring enables a coder to reorganize and edit to create a more robust, faster, better commented and easier to read script. When analysing a script for refactoring, redundnacies become more clear and duplicate pathways can be reduced. It is also a good way to understand a program. 
  
  
  
  **Advantages and Disadvantages of the original and refactored VBA script**
  
  When refactoring the VBA script, I realized why the previous code worked and was able to undestand it on a deeper level. The names of each variable were changed to make them more intuitive and the code became more viable for adding a larger data base. The biggest take away was that the formatting took a substantial hit on the run time of the code. By combining formatting into the script itself, rather than have it be a stand alone script, the UX was much improved but the formatting choices made by the guide ended up usurping any refactoring advantage that had been created by the non-formatted processed script. 
  
  
