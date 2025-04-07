Hello :)

In this file, I would like to present my knowledge of Excel in practical activities along with descriptions of what I do specifically. I operate on a database found on the Internet, concerning the sale of bicycles.

I have described each example with a header to make it easier for you to navigate.

1. Data cleaning + Grouping + Find and replace + Hiding personal data + absolute reference + converting text to numbers (SUBSTITUTE, LEFT, REPT, LEN)
2. Cell formatting + Sorting + Filtering 
3. Charts, financial and mathematics functions (SUMIFS, TRIM, INDEX + MATCH, SUMPRODUCT, SUM, UNIQUE, XLOOKUP)
4. Pivot tables
5. Financial calculations (Average, conversion to USD, profit, revenue gross/net, cost, average order value, unit cost, profit margin)

In the second file of this repository I have placed a csv file for download.

You can click the headers below to go directly to the specific chapters:

# Table of Contents
- [Table of Contents](#table-of-contents)
- [1. Data cleaning + Grouping + Find and replace + Hiding personal data + absolute reference + converting text to numbers](#1-data-cleaning--grouping--find-and-replace--hiding-personal-data--absolute-reference--converting-text-to-numbers)
- [2. Cell formatting + Sorting + Filtering](#2-cell-formatting--sorting--filtering)
- [3. Charts, financial and mathematics functions](#3-charts-financial-and-mathematics-functions)
- [4. Pivot tables](#4-pivot-tables)
- [5. Financial calculations](#5-financial-calculations)



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
Another thing worth doing before future analysis is proper grouping to smoothly move around the table. I suggest grouping and thus temporarily covering the columns: Product Subcategory, Product Name, Product Description, Product Size, Product Region, Product Color, Payment Method, Shipping Method.

![2 0](https://github.com/user-attachments/assets/87606dd1-21fb-4cb9-a73c-077ff7df5aef)

7. Personal data
- The last thing in preparing the data will be to cover the personal data with asterisks, for this purpose we will use a similar code as in the case of conversion to numbers, but we must determine that at least one letter of the name is visible
- The LEFT function will help us with this, which will take the first letter from the name and surname
Then the concatenation icon, i.e. "&", and then REPT from "repeat" so that the remaining letters of the name are also covered. The length is determined by the length of the characters in a given cell -1

![2 1](https://github.com/user-attachments/assets/87042fa1-a21e-4fba-9c6c-d9d600edac93)


[Table of Contents](#table-of-contents)
## **2. Cell formatting + Sorting + Filtering** 

1. Formatting opinions
- First of all, we can refer to opinions, so I suggest formatting the cells so that they have colors assigned to positive, negative and neutral opinions. To do this, select the cell with opinions, and after selecting cell formatting, select the appropriate colors assigned to the opinions.

![2 1 1](https://github.com/user-attachments/assets/f963074f-a7e3-49c1-b281-4b0c9d5dffac)

2. Data bars
We can then add chart columns to each cell to visually show which bikes had a price compared to others.

![2 1 2](https://github.com/user-attachments/assets/5fb502c9-7f5e-44db-b128-27d41763e2e7)

3. Highlight shipment
Now we can further highlight the canceled transactions.

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


[Table of Contents](#table-of-contents)
## **3. Charts, financial and mathematics functions**

1. Creating an additional table
To simplify our data and create charts, we can create a smaller helper table that will calculate the number of bikes sold to individual countries.
I suggest making such a table under the main table, so that on the left side, the rows are the names of the countries, and the columns are the types of bikes.

![3 1 1](https://github.com/user-attachments/assets/b14cf4e1-273e-4f1f-b686-892cc79d40d3)

2. Sumifs
Now we can use the sum if function to count how many times items were purchased for each country using specific cells.

![3 1 2](https://github.com/user-attachments/assets/3869ccbe-05a9-46e0-ae1e-bd7378aaebd7)

Using flash fill we can drag the function to the entire length of the table.
However, we can see that there must have been an error somewhere, the function is written correctly, yet Road Bikes shows zero. And in the previous steps we determined that the bikes were sent to Poland.

So the error must be somewhere in the cells we use for the equation. When we analyze each of them, we will see that in the country names there is a space before the country name, that's why the code does not work correctly.

![3 1 3](https://github.com/user-attachments/assets/16f94c0e-6493-4023-b190-609ee26ebb58)

3. TRIM
To get rid of spaces, we can type =TRIM(02), which will effectively remove unnecessary characters. Then we drag down to format all countries in this way and paste only the values ​​in place of the country names.

![3 1 4](https://github.com/user-attachments/assets/5430a885-8ea6-47f1-ab35-cd49be418b27)

Immediately after copying the countries without spaces, the number 3 appeared in Road Bikes, now we can drag the code to the remaining cells to fill the entire table. It is worth noting that hidden spaces are one of the most common difficulties when working with data.

![3 1 5](https://github.com/user-attachments/assets/bbf446e2-4f80-4e76-b616-776e2abc8721)

4. SUM
Now it's worth adding the SUM column, which will easily show us in which countries the most bikes are purchased
We can see that the USA is definitely in the lead, while Japan is in second place, and Germany in third.

![3 1 6](https://github.com/user-attachments/assets/70d7c40c-4ba7-49c3-8bad-ccca3a8e0318)

Now we can add the total amount of each type of bike sold at the bottom of the table. By selecting these cells and the cells with the bike types, we create a pie chart to show how the sales of each bike were distributed. 
By selecting the country column and TOTAL, we create a column chart to show more pictorially how the sales results are presented in each country.

![3 1 7](https://github.com/user-attachments/assets/59eb6b47-14da-410c-856c-0ecaed949c28)

5. UNIQUE + INDEX + MATCH
- Now, using the sum of the product, we can determine what share a particular type of model had in the total sales
To do this, we need to unwind the groups that were created earlier, so that we can see the names of the bikes. We can use the UNIQUE function to quickly copy the entire column and be sure that the records are not repeated. Additionally, if more bike models are added in the future, the list will update.

![3 1 8](https://github.com/user-attachments/assets/cc7fcab6-6507-488a-b21a-163b0eea9cca)

To add a price to individual bikes we can use INDEX + MATCH, thanks to this the price will always stick to a specific model

![3 1 9](https://github.com/user-attachments/assets/0ec24609-3ad4-4b7b-983c-ca875a3b9f05)

6. XLOOKUP
To accurately assign a price to each product, especially when the columns are grouped and covered, a good solution would be to use the XLOOKUP function
Thanks to this, we already know what the share of individual bike models is in the sold items, we can now calculate what their share is.

![3 2 0](https://github.com/user-attachments/assets/d9e53709-e681-471c-b325-3e1d12e7c9ca)

- I calculate the product sum - VERY IMPORTANT - not to use the sum, because then we will only calculate the sum of sold items, which is not the same as revenue
- We calculate the product sum from the price of the product and the quantity sold, which is divided by the price of the SPECIFIC product and its specific quantity, which gives us the division of each product by the total revenue.

![3 2 1](https://github.com/user-attachments/assets/6a577263-ff6c-47ae-a7b1-adb5eb361eed)

To visually show the proportions, we can use a column chart to show the differences.

![3 2 2](https://github.com/user-attachments/assets/0b024637-fb16-4876-8eaa-1f617a2b2b45)

7. Scatter Plot
Because the number of individual models sold is quite small = 1,2,3 subcategories work well to check the sales shares. Which are more detailed than general categories, but less detailed than individual bike models.

I suggest creating a similar table below as the one with product prices. Now, however, we will sum up all the bikes that belong to a given category using the SUMIF function. 

![3 2 3](https://github.com/user-attachments/assets/8e8d2fb0-74f7-44ca-a094-e9235b40a6d1)

We can calculate the number of sales in subcategories using the same SUMIF function, we change the lookup table to R, which contains the quantity. Thanks to this, we can create a Scatter Plot.

![3 2 4](https://github.com/user-attachments/assets/6b81681c-90b5-4c44-bb4d-b603ec9de9f1)

In the graph we can see that the largest quantities of bikes sold = 6 are in both the medium-expensive and the most expensive categories. However, the most visible purchasing pattern is the medium-expensive subcategories 4,000 - 6,500 euros with the quantity between 3 to 4 pieces.

[Table of Contents](#table-of-contents)
## **4. Pivot tables**

A pivot table is a great tool that provides us with important information in a very simple and fast way. The most important thing is to choose the right values ​​in the table.

A pivot table is created by selecting the entire table and choosing "Pivot table", it is best to create it in a new window.

![4 1 1](https://github.com/user-attachments/assets/6952417c-b05a-400f-ae16-f2039e0f7c8f)

For example, in the table below, I can quickly see how many and which L-size bikes were shipped to Australia over all the days that are counted in the main table

![4 1 2](https://github.com/user-attachments/assets/3cf91666-42e7-4ccc-b5bb-6976d53fc1fc)

In the same way, we can check how many bikes of a specific color were purchased in a given country. In this example, it is blue for Japan.

![4 1 3](https://github.com/user-attachments/assets/19f5967f-30d8-4feb-a553-17be7471eadc)

We can also calculate other values ​​than the sum, such as average, min, max, standard deviation, etc. In this example, it shows the average weight of the ordered bike, the biggest advantage of pivot tables is that we get additional information, such as the fact in which country which bike model is chosen.

![4 1 4](https://github.com/user-attachments/assets/2247d8c7-33b2-4074-9769-a7e8977aec96)

In pivot tables we can use the same field for several areas. In the example below the product category is assigned both as a row and as a value. In the columns is the type of payment, and the location is used as a filter, this gives us very important information in which country and which models of bikes were purchased using PayPal or Credit Card.

![4 1 5](https://github.com/user-attachments/assets/b3a6890a-c545-49d9-a10a-b5246aa171fa)

The pivot table refers to the main table, so if we add a new column or row to the main table, the pivot table will not update automatically.
To do this, select the pivot table and choose "Change Data Source", then increase the range of columns / rows from which data is to be retrieved, so that it actually matches the main table. After that, use the "Refresh All" button to update the information.

![4 1 6](https://github.com/user-attachments/assets/cd050c59-8697-4af9-8c4d-9d1901812838)

Pivot tables can be formatted in the same way as regular tables, we can also add a slicer to them, which will allow us to select values ​​in the table more dynamically, using a separate window and a timeline, which allows us to show records based on dates in a more graphic way. 
In this example, the table shows that from March 1 to March 16, products with prices of 1.700, 1.800 and 1.900 euros. 
Using percentage values, we can easily indicate how the values ​​are distributed among individual countries, thanks to which we know that over 66.67% of canceled transactions occurred on March 3.

![4 1 7](https://github.com/user-attachments/assets/54a901a2-dcd5-42af-9c2f-0f9ee79ab3be)


[Table of Contents](#table-of-contents)
## **5. Financial calculations (Average, conversion to USD, profit, revenue gros/net, cost, average order value, unit cost, profit margin)**

In this section, I would like to present additional functions that we can use to interpret the results

1. Average + USD + Data Validation
- We use AVERAGE to calculate the average order amount
In the column next to it I have placed the dollar rate from today, we can set Data Validation, so that if in the future someone enters a new dollar rate, to make sure that the amount entered will not be below 0.01


![5 1 1](https://github.com/user-attachments/assets/361061c6-4fba-4d4b-ad3a-622003e9915e)

![5 1 2](https://github.com/user-attachments/assets/12715366-8c39-4448-bb22-4717fc802008)

2. Gross / Net
In this example, I show how to calculate the net value of total revenue, i.e. net of tax, at a tax rate of 18%.

![5 1 3](https://github.com/user-attachments/assets/ab86da59-872d-412f-bfdc-3df6a84d0259)

3. Profit
One of the most important pieces of information is calculating the profit, below I present the gross and net profit, assuming that the costs amounted to 80% of the revenue and the tax is 18%

![5 1 4](https://github.com/user-attachments/assets/97a910de-75e5-473d-9193-fe897d813187)

4. AOV - Average Order Value
We can calculate the average order value by dividing the sales revenue by the number of orders, as in the example below

![5 1 5](https://github.com/user-attachments/assets/3c51b8c7-2450-49c4-aa8f-0e0a52cac412)

5. Unit cost
Unit cost is very useful for calculating how much money a company spends on producing one product. It is crucial when planning a budget and looking for opportunities to cut costs. The formula for this is the cost of goods sold / quantity of products sold.

![5 1 6](https://github.com/user-attachments/assets/29b61e59-ec8c-4c0a-88ee-af4a934b3872)

6. Profit Margin
The margin shows the company what the best prices are and helps monitor the efficiency and profitability of the given activities. The formula for margins is gross profit / revenue * 100. In this situation, it is not recommended to multiply by 100, because Excel will detect the percentage number when this form is set in a given cell.

![5 1 7](https://github.com/user-attachments/assets/ea711af0-717e-456a-bbe9-880ce67efcb3)


[Table of Contents](#table-of-contents)
























