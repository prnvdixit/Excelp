# Excelp
A software to automate some of the tasks "not" already included in Microsoft Excel. This can be used by teachers teaching finance (to check the solution sheets of students on-the-go), by researchers (to cumulate the data collected from different sources into one sheet), by Data Scientists (to create 3D Pie-Charts for better visualisation of data).

## Installations
```
sudo -H pip install openpyxl
```

## Results

1. Checking if two excel sheets are same or not - Can be used by teachers to match answer script with the solution sheets of different students. The output as shown in picture will show which cells are different and in which sheets exactly. The output will also tell if the student has given incomplete assignment (less number of sheets or less number of entries).


```
  - The Original Excel sheet (orig)
```
  ![alt text](https://github.com/prnvdixit/Excelp/blob/master/result_images/compare_orig.png)


```
  - The Excel sheet to compare (comp)
```
  ![alt text](https://github.com/prnvdixit/Excelp/blob/master/result_images/compare_comp.png)


```
  - The Final result (terminal-window)
```
  ![alt text](https://github.com/prnvdixit/Excelp/blob/master/result_images/compare_result.png)
  
  
2. If data had been collected by independent researchers (all stored in different excel sheets). Microsoft Excel supports merging different different Excel files to different sheets in same Excel file. This function enables cummulating data from different sheets into the main (first) sheet.


```
   - The Sheet1 of Excel file (main sheet)
```
  ![alt text](https://github.com/prnvdixit/Excelp/blob/master/result_images/orig_before_merge.png)
 
 
 ```
  - The Sheet2 of Excel file (to-be-merged sheet)
 ```
   ![alt text](https://github.com/prnvdixit/Excelp/blob/master/result_images/orig2_before_merge.png)
 
 
 ```
  - The Final Merged Excel file (main sheet)
 ```
  ![alt text](https://github.com/prnvdixit/Excelp/blob/master/result_images/orig_after_merge.png)
  
  
 3. To allow better data visualisation, the user may use the software to create 3D Pie-charts. 
 
 
 ```
    - The Final Excel file (after appending the created Pie-chart to the end of data)
 ```
   ![alt text](https://github.com/prnvdixit/Excelp/blob/master/result_images/pie_chart.png)
   
   
 ## Contributor

* **Pranav Dixit** - [Github](https://github.com/prnvdixit) - [Linkedin](https://www.linkedin.com/in/prnvdixit/)
