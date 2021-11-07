# Stock Analysis

## Overview of Project

### Purpose
We created a workbook that analyzed 2017 and 2018 stocks performance based on total daily volume and returns. Now we want to refactor the code that will return the same information and determine whether refactoring the code made the VBA scrip run faster.

## Results 

### Analysis of Stock Performance Between 2017 and 2018

![2017 Performance](https://user-images.githubusercontent.com/92401000/140660103-1ef887a4-092d-414c-ac0e-a615032e091f.PNG)

![2018 Performance](https://user-images.githubusercontent.com/92401000/140660105-0cc8a423-3c1d-4031-8f03-61aeb364d96a.PNG)

In 2017, 11 out of 12 green stocks had a positive return with the exception of TERP. In 2018, only 2 out of 12 green stocks (ENPH and RUN) had a positive return while the other 10 had a negative return. In 2018, we did not see any triple digit gains for green stocks compared to 2017, however for the two stocks with positive return in 2018 (ENPH and RUN), the total daily volume for those stocks increased at least 2x. 

### Analysis of Execution Times of Original and Refactored Scripts

![Original Run Time 2017](https://user-images.githubusercontent.com/92401000/140661013-9785f452-fb4b-47b8-8151-f89a7f5af6ee.PNG)

![Refactored Run Time 2017](https://user-images.githubusercontent.com/92401000/140661025-f3b02acb-4d59-46e3-b2a5-cd7f0ae0afde.PNG)

The original script ran in 0.832 seconds and the refactored script ran in 0.199 seconds for 2017. We saved 0.633 seconds with the refactored code.

![Original Run Time 2018](https://user-images.githubusercontent.com/92401000/140661033-389f4474-176d-417c-8f0e-8789ee689113.PNG)

![Refactored Run Time 2018](https://user-images.githubusercontent.com/92401000/140661036-270ed5f6-15ce-44dd-8c18-9f7214c29e3f.PNG)

The original script ran in 0.960 seconds and the refactored script ran in 0.211 seconds for 2018. We saved 0.749 seconds with the refactored code.

### Analysis of Code Differences of Original and Refactored Scripts

#### Original Script
![OG code](https://user-images.githubusercontent.com/92401000/140661286-ae31d9cc-9e46-4844-a3d7-26c8ce2e1eea.PNG)

#### Refactor Script

![Refact Code](https://user-images.githubusercontent.com/92401000/140661401-4911e299-34e8-47b5-931d-0e49f67315fe.PNG)

![tickerIndex](https://user-images.githubusercontent.com/92401000/140661525-a3fde0e0-a6f0-44ee-af53-a3a3acedaaac.PNG)

The main difference between the original script and refactored script was the original script used a nested for loop while the refactored script used multiple for loops. Also in the refactored script, we created a tickerIndex variable where it can be iterated throughout the script.

## Summary

### Advantages and Disadvantages of Refactoring Code

### Advantages and Disadvantages of the Original and Refactored VBA script 
