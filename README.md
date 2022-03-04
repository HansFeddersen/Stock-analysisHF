# Stock-analysisHF
Data Bootcamp Module 2: VBA
## Overview of Project

### The purpose of this project is to refactor a working code of an analysis for a data set with the objective of making it more efficient and to make it run faster. This a very usefull tool when working big with data set. In this particular case, the data set contains information about the open prices close prices and volume traded for 12 stocks in each day of the year 2017 and 2018. The written code, helps to obtain a summary of the return of each stock and the total volume traded for the selected year.


## Results

In order to see if the refactoring of the code worked, we have to take into consideration two main factors:

* **The outcomes for every stock are the same in the refactored code and the original:** That proves that the refactored code does not contain any mistakes. In this case, the outcomes are equal for both years, so the first factor is ok. The results for each year are as follows (image has been formated):

![This is an image](https://github.com/HansFeddersen/Stock-analysisHF/blob/main/Resources/more/StockOutcomes.png)

* **The refactored code did make the script run faster than the original code:** In this case, the refactored code ran faster than the original code for each of the years. Its important to keep in mind that everytime that the script is ran, the timer gives a different time, but times are always close to each other. For more cenrainty, the original script and the refactored one where ran one right after the other. The comparison between the timers of both scripts is as follows>

![This is an image](https://github.com/HansFeddersen/Stock-analysisHF/blob/main/Resources/more/Original_Refactored_Timer.png)

* **Code Comparison:** For a better understanding, it is necessary to present the refacored code with explanations:




## Questions

**What are the advantages or disadvantages of refactoring code?**

The main advantages of refactoring code are as follows:

* Makes the code run faster, so when working with bigger data sets its possible to make an analysis faster without affecting the outcome.
* To be refacotred, code must be clear, clean and in order, so a refactored code is easy to understand for people other than the author.
* In a refactored code it should be easier to identify mistakes.
* You can eliminate duplicates (if there were any in the original script).

The main disadvantages of refactoring code are as follows:

* If not done properly, can change the results and you will have a wrong analysis.
* If the code has lines that are extra (or duplicate) by mistake it could ran slower than the non refactored code, as one is trying to make the script more efficient could also make it less efficient.


**How do these pros and cons apply to refactoring the original VBA script?**

* If you have written a clean and clear code, everybody that understands the concept could understand the code without taking to much time. Also, if the author ever has to go back to his code, it will help him save a lot of time on trying to figure out how everything works. To avoid the cons of refactoring the code, it is neccesary to run the code periodically as one is making changes and compare it to the original result to avoid any mistakes and make sure that the refactoring is achieving its goal.
