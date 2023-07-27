# The Excel Process - From Dirty to Clean:                                       

*Google Sheets won't open files of these sizes. Excel is your only option to open the data. The instructions for Google Sheets are unfortunately pointless:*

* Download the most recent consecutive annual data here (which will be 12 files): https://divvy-tripdata.s3.amazonaws.com/index.html.
* Unzip all 12 files and place them in a specially named folder with CSV in the title for clarification.
* Place that folder in a location relevant to the material. The CSV files will be used for SQL, R, and Tableau later.
* <strong>AFTER</strong> we clean the data, <strong>THEN</strong> we will copy and convert all the files to XLSX inside a new folder with XLSX in the title. Copying now, as the instructions state, leaves you doing the same work twice: CSV files cannot save multiple sheets in one file.

*XLSX files can store more data and have a better compression algorithm, saving you space.*

### *If* the sheets have any missing data, remove the entire row
<details>
<summary> Cleaning the data </summary> 
  
  *This process is per situation, and normally stakeholders are involved in the decision on what to do with empty cells.*
  
  <ol>
    <li> Open the first file. Select all fields (including column names) by clicking and dragging over all columns or by clicking the utmost upper-left from the field section of the sheet. Aka, above row 1 and to the left of column A. </li>
    <li> After selecting all fields, press F5 or hold down CTRL+G until a "Go To" window pops up > Select "Special" > Click "Blanks" > Hit OK. This may take minutes to finish running. </li>
    <li> Once finished, scroll down until you see a highlighted cell or chunk of cells. Right-click when hovered over one and choose "Delete," then choose "Entire row" (you may get a warning; hit OK). This will take minutes, and your sheet may freeze; that's normal. </li>
    <li> Unfortunately, you must do all of these steps on your current sheet and follow sheets as many times as it takes until your results land you at the bottom of the sheet. The benefit of checking one too many times is you get the last row number. This information is useful.</li>
  </ol>
  <ul>
    <li> Normally you sort and filter each column depending on the data type looking for anomalies or any number of error values </li>
    <ul>
      <li> Currency: currency types that are out of range. </li>
      <li> Date: dates that are out of range. </li>
      <li> Number: numbers that are out of range. </li>
      <li> Percentage: percentages that are out of range. </li>
      <li> Text: letters or word lengths that are out of range. </li>
      <li> Time: times that are out of range. </li>
    </ul>
  </ul>
  
  *This data is much cleaner than in normal situations, but we will see one instance where it may be applicable.*
  
</details>  

### Now the data is clean and consistent, it's time to add new columns by using formulas
<details>
<summary> Adding ride_length </summary>
  
  *In truth, normally we would also touch base with the stakeholders to ask about removing certain ride_length duration ranges because high and low thresholds are anomalies, offer little insight, and skew most results, outside of rare instances.*
  
 <ol>
   <li> In your spreadsheet, create a column called "ride_length" in Column N, row 1. </li>
   <li>Calculate the length of each ride using the minus operator from columns C (started_at) and D (ended_at) Enter "=D2-C2" in cell N2. </li>
   <li> Your result will be a data type called "float". We need to change that into the time format of HH:MM:SS. </li>
   <li> Select N2 > right click > A window pop-up will appear select "Format Cells" (again Excel may freeze). </li>
   <li> While in the "Number" tab find "Category:" and change it to "Time" > Type: > "37:30:55" > hit OK. </li>
   <li> Select N2 > press CTRL+C > use macros to autofill the column (web search) or Select N2 > press CTRL+C > in N3, hold CTRL+SHIFT+down-arrow key > CTRL+V aka paste. </li>
   <li> Now we need to scroll down and find the last data-filled row + 1. Select that cell then hold the CTRL+SHIFT+down-arrow key again to delete the invalid entries (Use PAGE UP & DOWN to move smoothly when close to the bottom filled row).</li>
   
   ### Some months *might* have faulty "ride_length" data filled with ####### which SQL will not allow. As analyst doing our process step, let's investigate using "Sort" and deleting these rows. 
   
   ### Also, I suggest deleting all rows from "ride_length" with durations above 24:59:00. This is not a mandate, but depending on your SQL skill level, it may help.
   
   <li> Select <strong>ALL</strong> columns and click on the "Data" tab at the top of the sheet > click Sort > Sort by ride_length > Order Largest to Smallest. Any cells in "ride_length" filled with ####### need their whole row deleted (mind your header row). Then check from smallest to largest for any cells filled with #######. The reason the cells are messed up and filled with ####### is that the "start_at" and "end_at" column data are reading backwards. We don't have the authority to flip them, so our only means is to delete them. </li>
   
   * *Excel is a mess when sorting. It doesn't have the ability to use a primary key to sort all of the fields based on one column. If you forget to sort by <strong>all</strong> columns, your data will be wrong. Also, filtering is limited to 10,000 unique items; with files of this size, filtering for what we need to accomplish is useless.*
   
   <li> Now repeat these steps for all 12 sheets. </li>
</ol>
</details>

<details>
<summary> Adding day_of_week </summary>
  
<ol>
<li> In your spreadsheet, create a column called "day_of_week." in Column O, row 1. </li>
<li> In O2, enter "=WEEKDAY(C2,1)", 1 = Sunday, and 7 = Saturday. Later, if you prefer your column to have the actual weekday name, use "=TEXT(C2, "dddd")" but only after switching to XLSX. The only catch is =MODE() cannot use the TEXT data type. Workarounds include manually writing in the day name in your pivot tables or flipping the column formula when desired.</li>
<li> Select O2 > press CTRL+C > use macros to autofill the column (web search) or Select O2 > press CTRL+C > in O3, hold CTRL+SHIFT+down-arrow key > CTRL + V aka paste. </li> 
 <li> Again we need to scroll down and find the last data filled row + 1. Select that cell then hold CTRL+SHIFT+down-arrow key again to delete the invalid entries (use PAGE UP and PAGE DOWN to move smoothly when close to the bottom filled row). </li>
 <li> Repeat these steps for all 12 sheets, and make sure to save your work. We're done with the CSV files until SQL and R. </li>
 </ol>
</details>

<details>
  <summary> Combining all files into one sheet </summary> 
  
  *<strong>NOW</strong> We are going to copy and convert all the files to XLSX inside a new folder with XLSX in the title.*
  
<ol>
<li> Open the first clean CSV file. </li>
<li> File > Save As > Browse > Your XLSX folder location > Save as type: Excel Workbook. Do this for all 12. </li>
<li> We are going to merge the 11 other sheets into the first sheet by creating new tabs at the bottom by using Power Query (Google search) or simply copying and pasting each sheet with CTRL+A > CTRL+C > and then pasting into a new tab in the first sheet with CTRL+V. Your first sheet will have 12 tabs when finished. </li> 
<li> Do this for all 12. Be mindful to keep your sheet tab names consistent if you're copying and pasting. They won't auto-populate. </li> 
</ol>
  
*Notice all your file sizes are smaller now and you now have a mega file too :clap:.*
</details>  


# The Excel Process - Time to Analyze:
<details>
  <summary> Run a few calculations </summary> 
  
  *Switch to the XLSX mega file now. Run a few calculations in two tabs of opposite seasons to get a better sense of the data layout.* 
  
<ol>
<li> Calculate the mean of ride_length: in cell Q2, type =AVERAGE(N:N), then format to time just like when we made column N "ride_length". Then make a header in Q1 so you remember what your result represents. </li>
<li> Calculate the max ride_length: in cell Q5, enter =MAX(N:N), then format to time again. Then make a header in Q4 so you remember what your result represents. </li>
<li> Calculate the mode for day_of_week: in cell Q8 enter =MODE(O:O). Then make a header in Q7 so you remember what your result represents. </li>  
</ol>
</details>  

## Pivot tables and Graphs
<details>
<summary> Calculate the average ride_length for members and casual riders </summary>
  
  *Create these 3 pivot tables and graphs in both pre-selected opposite season tabs used earlier.* 
  
<ol>
<li> In cell Q11 click "Insert" on the top tab > Click "PivotTable" > select columns M and N together > Existing Worksheet then OK. </li>
<li> Drag member_casual in the Rows area and ride_length in the Values area > Left-click it and choose "Value Field Settings" change Count to Average. </li>
  
  * (blank) auto populates inside your pivot table, this is normal. Remove (blank) by clicking on cell Q11
  
<li> Now that you have your first pivot table, it is time to format R12-R14 just like column N "ride_length" to time. </li>
<li> The last step is to graph it. Click Q11 > at the top of Excel, click "Insert" > "Recommended Charts" > "Pie". </li>
<li> Place its upper-left corner in Q15. Use whatever chart you like. I just find pie to be the best for this table. </li>
<li> Select the chart and click on "chart styles". Pick whatever variation you like. Shrink the graph to your preference. </li>
</ol>
</details>

<details>
<summary> Calculate the average ride_length for users by day_of_week </summary>
<ol>
<li> In cell Q29, click "Insert" on the top tab > Click "PivotTable" > select columns M, N and O together > Existing Worksheet then OK. </li>
<li> Drag member_casual in the Rows area and ride_length in the Values area > Left-click it and choose "Value Field Settings" change Count to Average. Finally, drag day_of_week into the Columns area. </li>
  
 * (blank) auto populates inside your pivot table; this is normal. Remove (blank) by clicking on cell Q31
<li> Now it is time to format R31-Y33 just like column N "ride_length" to time. </li>
<li> Time to graph it. Click Q29 > at the top of Excel, click "Insert" > "Recommended Charts" > "Column." </li>
<li> Place its upper-left corner in Q34. Use whatever chart you like. I just find columns to be the best for this table. </li>
<li> It is recommended you change the day-of-week color palette for 1 & 2 because they match the first graph. </li> 
<li> Click into the new graph, then click a bar. Right-click it once selected and select "fill". Stretch the graph to column Y. </li>
</ol>
</details>

<details>
<summary> Calculate the number of rides for users by day_of_week </summary>
  
*This one is a little tricky.* 
<ol>
<li> In cell T3, click "Insert" on the top tab > Click "PivotTable" > select columns A, M and O together > Existing worksheet, then OK. </li>
<li> Drag ride_id into the Value area and make sure its "Value Field Setting is set to Count. Then drag member_casual and day_of_week to the Rows area. </li>
  
   * (blank) auto-populates inside your pivot table; this is normal. Remove (blank) by clicking on cell T3
 
  <details>
  <summary> Solution </summary>
    <ol>
<li> To get the pivot table to form, you must select all columns.</li>
<li> To get the pivot table to be sorted by day instead of member type, day_of_week must be loaded into the Rows area first. </li>
      </ol>
    </details>
<li> It's time to graph it. Click T3 > At the top of Excel, click "Insert" > "Recommended Charts" > "Column." </li>
<li> Place its upper-left corner in V3. Use whatever chart you like. I just find columns to be the best for this table too. </li>
<li> It is recommended that you change the member color to the same orange as the first graph. </li>
<li> It is also recommended to stretch the graph to column Z and row 25. </li>
<li> Also, you can change day_of_week to text format now. </li>
</ol>
</details>

## Summary 
* There are some visual edits you can make.
* Take your two seasonal analysis sheets and merge the pivot tables and graphs into one new sheet.
* Finally, take a screenshot of it, Now your Excel summary visuals are complete!

Due to the amount of content, the summary is too large to post on GitHub; a .PNG file has to do.

<details>

# In The End
Per stakeholder request: 
* We downloaded, cleaned, sorted, and converted the CSV files with no error values allowed
* We added new columns by combining data using formulas
* We created multiple Pivot tables and Graphs based on stakeholder requests
* We created a summary and highlighted two opposing seasonal months of data
