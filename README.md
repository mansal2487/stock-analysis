# Stock Analysis

## Overview of Project

I wrote VBA code to that allows someone to gain insight on a stocks performance during different years. The code did what it's supposed to do but I wanted to make it run faster. To do this, I refactored the code to only loop through the stock data once instead of multiple times.

## Results

### Data Analysis

Below are two images that show help to understand how some stocks performed in 2017 and 2018. The data was gathered using the refactored VBA code.

![alt text](https://github.com/mansal2487/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

![alt text](https://github.com/mansal2487/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

As you can clearly see, these stocks collectively had a higher return rate in 2017 than they did in 2018. A Misnky Moment may be the cause of this. Named after economist Hyman Minsky, a Minsky Moment refers to the idea that long enough periods of bullish speculation will eventually lead to crisis.

### Code Analysis

Below are snippets of the code before refactoring and after refactoring, repspectively.

![alt text](https://github.com/mansal2487/stock-analysis/blob/main/Before_refactor.PNG)

As you'll see above, the code loops through the hundreds of rows of stock data 12 times before terminating. This takes over one second to run. That may not sound like a lot, but when you write an extremely long script, managing this becomes very important. 

![alt text](https://github.com/mansal2487/stock-analysis/blob/main/After_refactor.PNG)

Above is a picture of part of the refactored code. Instead of looping through the stock data multiple times, we gather all of the information in one loop by using a variable called the "tickerIndex" that represents the stock that we're looking for. It increases by one when the loop gets to the last record of data for that stock which tells the loop that the next chunk of data it will find and store is going to be for the stock that corresponds to this new number. This part of the code takes about three tenths of a second to run or about three times faster than the code before refactoring.

## Summary

### Refactoring in General
Refactoring code can make the code more readable, faster, extensible for use with other functions, autonomous, and use less memory. Now, all of this sounds great and might make you think, why wouldn't you refactor all code until it's "perfect"? The answer is because refactoring code can take a long time and get you flat out confused. Unless you are trying to achieve something previously mentioned and it's necessary, for the purpose of learning, or both, I think it's best to steer clear of it.

### Refactoring the VBA

Refactoring the VBA code was both necessary and for the purpose of learning. It made the code execute a lot faster and also made it autonomous in terms of not having to know exactly how many unique stocks are apart of the data set. The disadvantages are that it took me quite some time to complete and the new code might be confusing to someone who is new to code versus someone with more experience.



