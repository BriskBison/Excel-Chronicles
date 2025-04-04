Hello :)

In this file I would like to present my knowledge of Excel in practical activities along with descriptions of what I do specifically. I operate on a database found on the Internet, concerning the sale of bicycles.

I have described each example with a header to make it easier for you to navigate, below is the table of contents

1. Data cleaning + Grouping + Find and replace + Hiding personal data + absolute reference + converting text to numbers
2. Cell formatting + Sorting + Filtering
3. Pivot tables and operations on them
4. VLookup / HLookup / Xlookup
6. Financial functions and others

The remaining files in this repository are specific records of the Excel file, for each topic, starting with the raw file, which is the file before any changes.


# **1. Data cleaning + Grouping + Find and replace + Hiding personal data + absolute reference + converting text to numbers** H1

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
- It is very important to paste only the values, because without this code the formula would be replaced and then what we did in Substitue would not work correctly
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


# **2. Cell formatting + Sorting + Filtering** H1
















