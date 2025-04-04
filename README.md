Hello :)

In this file, I would like to present my knowledge of Excel in practical activities along with descriptions of what I do specifically. I operate on a database found on the Internet, concerning the sale of bicycles.

I have described each example with a header to make it easier for you to navigate.

1. Data cleaning + Grouping + Find and replace + Hiding personal data + absolute reference + converting text to numbers (SUBSTITUTE, LEFT, REPT, LEN)
2. Cell formatting + Sorting + Filtering 
3. Charts, financial and mathematics functions (SUM, COUNTIF, INDEX + MATCH, SUMPRODUCT)
4. VLookup / HLookup / Xlookup
6. Financial functions and others


You can click headers below to go directly to the specific chapters:

- [1. Data cleaning + Grouping + Find and replace + Hiding personal data + absolute reference + converting text to numbers](#1.Data-cleaning-Grouping-Find-and-replace-Hiding-personal-data-absolute-reference-converting-text-to-numbers)
- [2. Cell formatting](#2.-Cell-formatting)
- [3. Charts, financial and mathematics functions](#3.Charts-financial-and-mathematics-functions)

The remaining files in this repository are specific records of the Excel file, for each topic, starting with the raw file, which is the file before any changes.


## **1. Data cleaning + Grouping + Find and replace + Hiding personal data + absolute reference + converting text to numbers** 

1. Duplicates
First select everything CTRL + A and remove duplicates

![1 1](https://github.com/user-attachments/assets/b0f74b08-f69c-4931-9ed2-54d5026344cc)

2. Empty cells
- Using "Find and replace" we search for cells to replace
- We can do the same for "spaces", "N/A" errors or any other value that may affect the data analysis

![1 2](https://github.com/user-attachments/assets/7198182a-8463-447a-b8ac-d849d7ec8997)

3. Autofit
- For a better view, we can adjust the automatic width and height of the cells, so that the data is more readable and the ####### signs that indicate

![1 3](https://github.com/user-attachments/assets/68e29dc3-a033-4807-8f71-7adab2b842e1)

4. Changing text to numbers
- The next thing we should do is add the Euro sign to Product Price, but when we try to do this it turns out that nothing changes. This is because this data is entered as text, not as numbers. One way to convert them to numbers is to change the dot to a comma and then multiply everything by the number "1"
- In this situation, we can use the Substitute function, mark the range of the entire product price column, change the dots in it to a comma, and then multiply everything by 1
- Then we can use flash fill, but remember to add the absolute value ($ symbol) to the cell where our number 1 is, so that it does not move during filling
- The next step is to copy the special "values ​​only" to the Product Price column
- It is very important to paste only the values, because without this code, the formula would be replaced, and then what we did in Substitute would not work correctly
- After pasting, we can easily select the entire cell and change the values ​​to currency. We should do the same with the Product Weight and Order Total columns

![1 4](https://github.com/user-attachments/assets/e9189305-924e-4f81-8030-660da76e3f1d)

![1 5](https://github.com/user-attachments/assets/c968bf74-9a12-4699-aa9e-1c0bafb4576d)

![1 6](https://github.com/user-attachments/assets/fa239601-59d9-4773-b448-5025f7c3ca2c)

![1 7](https://github.com/user-attachments/assets/8cf6283e-b2e3-4d3b-ab60-aec657a49ec4)


5. Enlarge headers
- For better readability, it is worth bolding the headers, increasing their font and bolding the border, changing the color of the header cells and Product ID for easier searching, which is especially important when working with SQL
- After that, you should re-select everything using ctrl + A and increase the cell width to auto-fit

![1 8](https://github.com/user-attachments/assets/a74d8675-8357-4da4-abfb-566d2e2812e0)

![1 9](https://github.com/user-attachments/assets/fe08b323-af79-446c-b62b-99002ab73da3)

6. Grouping
Another thing worth doing before future analysis is proper grouping to smoothly move around the table. I suggest grouping and thus temporarily covering the columns: Product Subcategory, Product Name, Product Description, Product Size, Product Region, Product Color, Payment Method, Shipping Method

![2 0](https://github.com/user-attachments/assets/87606dd1-21fb-4cb9-a73c-077ff7df5aef)

7. Personal data
- The last thing in preparing the data will be to cover the personal data with asterisks, for this purpose we will use a similar code as in the case of conversion to numbers, but we must determine that at least one letter of the name is visible.
- The LEFT function will help us with this, which will take the first letter from the name and surname
Then the concatenation icon, i.e. "&", and then REPT from "repeat" so that the remaining letters of the name are also covered. The length is determined by the length of the characters in a given cell -1

![2 1](https://github.com/user-attachments/assets/87042fa1-a21e-4fba-9c6c-d9d600edac93)


## **2. Cell formatting** 

1. Formatting opinions
- First of all, we can refer to opinions, so I suggest formatting the cells so that they have colors assigned to positive, negative and neutral opinions. To do this, select the cell with opinions, and after selecting cell formatting, select the appropriate colors assigned to the opinions.

![2 1 1](https://github.com/user-attachments/assets/f963074f-a7e3-49c1-b281-4b0c9d5dffac)

2. Data bars
We can then add chart columns to each cell to visually show which bikes had a price compared to others

![2 1 2](https://github.com/user-attachments/assets/5fb502c9-7f5e-44db-b128-27d41763e2e7)

3. Highlight shipment
Now we can further highlight the canceled transactions

![2 1 3](https://github.com/user-attachments/assets/61314563-d5d9-4692-a5f8-28cca6e26d59)

4. Sorting
- Now sorting will be easier, because we can do it based on color, for this purpose, we select the Order Status column, right-click on sorting and choose color and descending order
- Now we can add another sorting condition. I suggest choosing the color of the feedback to check if the cancellation of the shipment was associated with a negative opinion
Adding colors and mini charts is very helpful, thanks to this we can already draw preliminary conclusions from a distance. As you can see, negative opinions were issued to items that were sent correctly. Therefore, other factors must have influenced the cancellation of the order as well as negative opinions. In addition, thanks to the columns assigned to the prices of bicycles, we can see that there is no direct connection between price and opinions / cancellation, because in both categories there were items from every price range.

![2 1 4](https://github.com/user-attachments/assets/5aa64d8c-4751-4647-b639-83cacc904b93)

![2 1 5](https://github.com/user-attachments/assets/185e9bc8-7c0e-42d8-bebe-3978eadf647a)

![2 1 6](https://github.com/user-attachments/assets/f91bc146-8c7b-425a-906b-ea8a1a41f26d)

5. Filtering
- Zooming in, we can see that only Road Bikes and BMX are at the top of the canceled bikes, we can use filtering to check if it's just a coincidence, the fact that Road Bikes sell the most, or if it's actually them that's the problem
- To do this, we'll select Product Category and select the filter and Road Bikes
- I also suggest grouping the Product Weight, Product Stock, Order ID, Customer ID and Customer Name columns to make the view more readable
- We can see that Road Bikes actually contains both canceled and positive ones, it would certainly be worth taking a closer look at what is the reason for the cancellation of Road Bikes, in a moment we'll create a chart to see how the number of Road Bikes is distributed compared to other bikes.
- Firstly I suggest sorting the bikes by price to see if price matters here

![2 1 7](https://github.com/user-attachments/assets/64da3aaa-1940-4aab-9f6a-cc910c741674)

- Now select the filter and Greater than or equal to 2800.00 to extract only the most expensive bikes

![2 1 8](https://github.com/user-attachments/assets/28ee4ce9-e727-4a01-aec8-f2d3ec51d85a)

- For this purpose, it will be very useful to extract the Top 10 from the order total column or simply sort descending by order amount. Sorting by number of orders and price ultimately showed that all cancellations, negative and neutral opinions mainly concerned large orders, especially those with 3 ordered bikes.

![2 1 9](https://github.com/user-attachments/assets/18b3b9a7-edb2-47a5-b7f8-0b8dda27be2a)

![2 2 0](https://github.com/user-attachments/assets/5f0e6247-37b5-4ec3-be93-b4d17fe9f013)


## **3. Charts, financial and mathematics functions**


























