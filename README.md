# Stock-Analysis

### Overview of the Project
The purpose of this project was to create a VBA analysis of a group of stocks and interpret what stocks in 2017 and 2018 were more successful. This could educate the public in terms of where they can put their money towards without such a high risk of losing it. With VBA, we can create macros that can automate repetitive word and data processing functions which we tend to use when using Excel. This application aid in creating spreadsheets as well as calculating a large amount of data such as this stock analysis.

### Results
Looking at the results, we see that in 2017 there were more stocks that had an increase in price between their open and close dates compared to 2018. We can see this not only by looking at the numbers in the "Return" column, but by the color coding of green as an increase and red as a decrease in price.

![2017](https://user-images.githubusercontent.com/57331058/132961694-4bf078a9-7444-415c-a3e3-120e96083613.PNG)

![2018](https://user-images.githubusercontent.com/57331058/132961700-117ef5a0-5644-4feb-9c5d-ca4793e61ef5.PNG)

In terms of the execution times between the original code and the refactored code, we see an drastic decrease in time from ~1.1 to ~1.3 seconds to ~.17 to ~.18 seconds. Below are the times of the refactored script of 2017 and 2018. Cleaning up the data, or refactoring, gave us an improved performance by not only shortening the code but by also reducing the complexity of such code.

![2017 Running Time](https://user-images.githubusercontent.com/57331058/132961773-ce12bb83-a5b2-491e-97a0-9ff77700a906.PNG)
![2018 Running Time](https://user-images.githubusercontent.com/57331058/132961777-3536b6bb-91cb-423b-879b-2dbf8151e8e9.PNG)

### Summary
Some advantages that come with refactoring code is that it includes code readability and reduces time as well as complexity. Refactoring cleans the code and improves performance. Not only does refactoring make it short and clean, but this kind of modification helps with finding bugs and fixing them a lot simpler and makes the software easier to understand.

One disadvantage or difficulty that comes with refactoring code is acknowledging what can be refactored and what has to stay the way it is; sometimes it becomes risky trying to refactor code when the application is big.

Applying it to this code, we have clearly seen that it improved performance by reducing the time tremendously and it's a lot easier to read and understand I'm doing with the stocks in terms of grouping each stock, analyzing each return by the open and close dates, and creating a color coded spreadsheet based on the positive/negative return and the year the user wants to see. Also, having shorter and simpler code helped me fix mistakes a lot easier than if I were to leave it how it was originally and trying to find the bug.
For my code I didn't necessarily get into much conflict, but one difficulty I can see is for a user to want to refactor it as much as possible and that can easily unnecessarily take a lot of time as well as it may not even function the same and that's where we can run into all kinds of bugs not knowing where the root of the problem is.

