# Refactoring VBA Stock Analysis Project

## **Goal**

For this assignment I was tasked with refactoring a VBA script with the goal of improving it's run time by simplifying the code. This task was a response to the customers intent to expand the amount of data processed by the macro. The original VBA script needed to re-loop through all of the data multiple time in order to process the output, which would have slowed down the runtime of the code as the amount of data increased.

## **Solution**

The main solution for this refactorization was utilizing loop functions execute multiple functions on the first run-through of the stock data, rather than the original code which re-looped the entire dataset before reprocessing the macro for each output.

## **Results**

In the end, our refactored code saw a significant increase in the scripts runtime (see attached runTimeTest.png). The refactored code was able to process the 2017 data .1796875 seconds faster while 2018 data saw a .5375976 second improvement.


## **Summary**

### **Pros and Cons of Refactoring Code in General**

### ○ Pros

The main goal of this assignment was to attempt to increase the speed at which the VBA can run. VBA overall is a good gateway to entry level data analysis, however these entry level macros are normally limited to very light analytic where the dataset is small. As that the size of the dataset increases, bad code can begin to bog down the users system memory. Refactoring is essentially cutting out the inefficiencies and streamlining the code as much as possible.

### ○ Cons

Cost of Resources: Although streamlining is normally a good thing, the project owner need to first weigh the potential gain vs the resources cost to refactor said code. If the person refactoring the code did not write the initial code, the resources spent on paying them to read, learn, and edit each line of code might ultimately outweigh the milliseconds saved on each macro run.

Potential Introduction of Errors: The project owner must also weigh the potential cost of unforeseen bugs introduced with the new code. If the code complex and there are high stakes riding on the code runs successfully every single time, then the project owner must again weight the cost of millisecond gained vs risk of downtime.
		
## **Pros and Cons of Our Refactored Code**

### **○ Advantages**
As previously mentioned, speed and the ability to process larger and larger datasets was the main driver for refactoring. When testing the refactored code, we saw a significant improvement in the run times. The 2017 data saw a 24% increase in runtime speed while the 2018 data saw a whopping 74% increase.

### **○ Disadvantages**
The major disadvantage to our refactored code is that it is still limited to the pre-defined number of stocks tickers added to the array. If the user eventually wanted to broaden their analyses, they would need to individually add the tickers in the VBA coding.
			
A second disadvantage in our code is that it uses a referenced based IF Functions to determine if the row it's reading is the first stocked traded in the year vs the last. If the data was scrambled out of order, this IF Function would fail to pull the correct stock prices for our Return calculation.
