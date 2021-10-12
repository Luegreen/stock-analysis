 # Stock-analysis
 ### VBA scripting, writing repeatable code for Excel, refactoring that code for speed and production. 
 
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
 
On first analysis, the refactored script ran slower than my original script by .05 seconds for 2017 and .03 seconds for 2018. This was discouraging until I realized that the refactored script was also formatting the final graph. My original script had no formatting. I built the formatting script separately and added it into the GUI as a second button option (run the script by a button to analyse the stocks, then push another button to format your graph). So, to get an apples to apples comparison, I realized that I had to equalize the scripts. I then took the formatting out of the refactored script to time it's run, and also I added formatting to the orignial script, and timed it's run. 

 The script runs any year with respect to the user's needs:
 <img width="1193" alt="Screen Shot 2021-10-12 at 2 57 14 PM" src="https://user-images.githubusercontent.com/14239715/137013475-34c55322-ad50-484a-856f-f9b11d605ec8.png">
<img width="1191" alt="Screen Shot 2021-10-12 at 2 57 28 PM" src="https://user-images.githubusercontent.com/14239715/137013501-b82dd4a9-9fbb-451b-b805-abf514738656.png">
<img width="713" alt="Screen Shot 2021-10-12 at 2 48 31 PM" src="https://user-images.githubusercontent.com/14239715/137012624-7715e1ba-5f9c-403d-a9f6-52ecbaa96bc2.png">
<img width="718" alt="Screen Shot 2021-10-12 at 2 48 42 PM" src="https://user-images.githubusercontent.com/14239715/137012627-80013f34-b09a-4f37-9248-21ee50402f9f.png">
<img width="668" alt="Screen Shot 2021-10-12 at 2 49 11 PM" src="https://user-images.githubusercontent.com/14239715/137012630-f8f6f24d-ba24-4c1a-a8c6-2759c2f3bf0a.png">
<img width="1225" alt="Screen Shot 2021-10-12 at 2 50 04 PM" src="https://user-images.githubusercontent.com/14239715/137012635-33adf7db-a4fb-4670-9be9-8d5718ffd5fa.png">




The results are as follows:
 
 
 <img width="330" alt="Screen Shot 2021-04-25 at 9 16 20 PM" src="https://user-images.githubusercontent.com/14239715/116016957-41482e80-a60c-11eb-9c45-956a42b4c7cd.png">
 
 In this chart you can see that while refactoring did save time for the pre-formatted script (.05 seconds for 2017 and .07. seconds for 2018), the formatted scripts, both before and after refactoring had the same run time, therefor refactoring saved no time for formatted scripts. 



 
 ### Summary
 
  **Advantages and Disadvantages of refactoring code in general**
  
  Refactoring enables a coder to reorganize and edit to create a more robust, faster, better commented and easier to read script. When analysing a script for refactoring, redundnacies become more clear and duplicate pathways can be reduced. It is also a good way to understand a program. 
  
  A disadvantage to refactoring is that you may spend a fair amount of time working with a code and not actually make it better. There are styles of coding and in refactoring you may make the code complient to your preference but not actually make it more robest. In some ways you will only gain a knowledge of the workings of the code... and I suppose you may make the code worse by having it take longer to process. 
  
  
  
  **Advantages and Disadvantages of the original and refactored VBA script**
  
  When refactoring the VBA script, I realized why the previous code worked and was able to understand it on a deeper level. The names of each variable were changed to make them more intuitive and the code became more viable for adding a larger data base. The biggest take away was that the formatting took a substantial hit on the run time of the code. By combining formatting into the script itself, rather than have it be a stand alone script, the UX was much improved but the formatting choices made by the guide ended up usurping any refactoring advantage that had been created by the non-formatted processed script. 
  
  If I had more time and deeper skills, I'd like to refactor the piece of code that now had the ticker indexes listed by hard code. I could make the script just select the 'next' ticker code and I'd be curious to see if that ran faster than the hard code. Even if it wasn't faster, it would reduce potential human error and also make it more useful as a code block for analysis of other scripts. 
  
  
