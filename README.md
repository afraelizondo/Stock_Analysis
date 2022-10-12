# Stock Analysis


## Overview of project

### VBA of Wall Street

During the current project we have been working together with Steve, who just finished his Financial Degree. This analysis was performed with the main goal of helping his parents to evaluate a possible investment into a New Energy Corporation called DAQO, a company that works with silicon wafers for solar panels, although this might seem like a good idea at the moment, Steve wants to show them other alternatives so they don't invest all their money in the same company, this way their funds can be more diversified in other green energy stocks. 

To make this analysis more concise, we decided to use Visual Basic for Applications, a human-readable and editable programming code which interacts with Excel, making calculations and performing analysis. One of the most significant advantages of this is that the tasks on the spreadsheet will be executed in exactly same way as long as we decide it and will be so much faster than do them manually. 



## Results 

The following chart shows the results of All Stocks macro that was runned for the 2018 year. As we can see there are few statements we can TALK about:
- There are 10 negative returns, which represents a 83.33% out of the total outcomes.
- The worst was "DQ" with a -62.6%.
- We have 2 positive returns, the highest was "RUN" with 84.0% and then "ENPH" with 81.9%.

<img width="224" alt="AllStocks" src="https://user-images.githubusercontent.com/113856917/195257691-d513da59-a54d-464b-a600-6609fb7ca30a.png">

Based on that, we can conclude there are more chances of having a negative return in the company, so is not convenient to invest all the money here, because there is a huge possibility that this will make more losses than gains. 

As terms of chart outputs, we got the same ones using the original and refactored script. 

In the execution statements we considerably improve, as is shown in Image 1 (original) and Image 2 (refactored). 

<img width="239" alt="VBA_Challenge 2018  " src="https://user-images.githubusercontent.com/113856917/195260869-a2cb38bc-b49a-455c-b961-3ac20e548058.png">

***Image 1***

<img width="243" alt="VBA_Challenge 2018  " src="https://user-images.githubusercontent.com/113856917/195260936-6442307b-25c7-49c7-b622-5dcea5c8a6dc.png">

***Image 2***



## Summary

In the previous project we made a process called Code Refactoring, which means the restructuration of the code without changing the original function or external behavior. This brings a lot of benefits to the coding process, because it improves the productivity and implies techniques to make the code faster to execute. We can list some of the advantages of the mentioned process:
1. Helps to find bugs.
2. Improves software design and makes it easier to understand. 
3. Avoids code duplications.
4. Removes excessive or confusing code. 
5. Clean code with shorter and self-contained methods. 
6. Better testability. 

Even though it has many advantages, there are some disadvantages as well: 
1. The possibility of introduction of new bugs and errors. 
2. No clear definition of "neat code".
3. In team work, it requires higher coordination than expected.
4. Sometimes the benefit is not self-evident due to the fact that the functionality stays exactly the same. 

These pros were applied to our VBA macro, and we were able to change the code and make it run faster than the first time. 

Also, to check if the current row was the first row with selected ticker and the last row with selected ticker using the *if* function, instead of only checking the starting and ending prices. We were able to make this in less steps than the first time. 

And finally, we made one *for loop* with outputs of ticker, total daily volume and return, instead of making 3 different ones. 

## ***Bibliography***

Refactoring: how to improve source code. (s/f). IONOS Digital Guide. 10/22/2022, https://www.ionos.com/digitalguide/websites/web-development/what-is-refactoring/

Veerraju, R. P. S. P., Rao, A. S., Murali, G., Paruya, S., Kar, S., & Roy, S. (2010). Refactoring and Its Benefits. 10/22/2022, https://aip.scitation.org/doi/abs/10.1063/1.3516393?journalCode=apc#:~:text=Refactoring%20improves%20the%20design%20of,the%20implementation%20when%20not%20refactoring.

