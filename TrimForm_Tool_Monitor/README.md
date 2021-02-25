## TrimForm Tool Monitor 
In the factory, there are machines that trims and forms the IC chips. Within these machines are tools used to achieve trimming and forming. As with all tools, the tools in the machines wear over time, and after a certain time, they would start producing defected lots. Therefore, engineers would need to replace these tools with new ones. However, there were no methods of detecting worn tools before defected lots were produced, meaning engineers could only replace the tools after the tools produce defected lots, which would then result in thousands of defected and wasted chips.

In the past, to find out whether the lots are defected, engineers need to manually read millions rows of data consisting of the dimensions of chips. The engineers would compare the dimensions to a set of thresholds to determine if it is time to replace the tools (for example, if the tip to tip length of the chips is higher than what the threshold states, that means the tool is worn and needs to be replaced). This task becomes tedious and arduous as the data is too large for any engineers to read, and it doesn't achieve the goal of replacing tools before defected lots are produced. Thus, the goal of this project is to alert engineers to replace the tools before defected lots are produced.

The Python script I wrote creates a user interface using PySimpleGUI, as shown below. The general function of the program is to read in the aforementioned data, insert tooling number via SQL (the raw data does not contain tooling number, therefore we would not know which tool produces which lot), and find the max values of each time period. Afterwards, user can choose to either see all products' dimensions' trends over time or one product at a time. 
<br/>
<br/>
<img src="pics/User Interface.PNG" width="75%">
<br/>
<br/>
Here I will detail each button's function.
- The "Update SO TSSOP" and "Update QFP" buttons updates and accumulates the data (and also saves data to excel files).
- The two "Find Max Value" buttons finds max values of each time period.
- "Clear Output Window" is for clearing the texts shown in the debugger window.
- The drop-down menus allow user to choose which product, which category (dimension), which tooling number, and which product pin number to see. Data Mode drop-down allows user to choose to look at actual numbers or percentages.
- The update buttons beside the drop-down menus updates the selection of each corresponding menus.
- "Generate Full Plot" shows the full plot of selected tooling.
- "Show Full Raw Data" shows the full data.
- "Full Average Raw Data" calculates the averages of values of each time period and then show it on screen.
- Users can input desired timeframe to show in days, for example, users can choose to look at the last 7 days of data.
- Users can input desired threshold for determining whether to replace the tooling or not. I'll detail the function of threshold more down below along with pictures.
- "Generate Plots and Tables" generates plots and tables corresponding to the input timeframe and threshold.
- "Generate Raw Data" generates raw data corresponding to the input timeframe and threshold.
- "Average Raw Data" calculates the averages of values of each time period and then show it on screen corresponding to the input timeframe and threshold.

Below is an example of using "Generate Plots and Tables" with a selected 80% threshold and 14 days timeframe on SOIC chip.
<br/>
<br/>
<img src="pics/Example.PNG" width="75%">
<br/>
<br/>
Each plot title shows the selected product, its corresponding tooling number, and pin number. Red lines represent over 50% of points during the last 14 days are higher than threshold, and blue lines represent less than 50% of points during the last 14 days are lower than threshold. These plots clearly tells user which tool needs to be replaced.

Below is another example, using the same button function, but this time on a single product.
<br/>
<br/>
<img src="pics/Example2.PNG" width="75%">
<br/>
<br/>
This plot shows the limit in red (100%), chosen threshold in green (80%), a regresion line in purple, and the line plot of Tip to Tip percentage. If more than 50% of the points are higher than the threshold, the plot will show an alert title, indicating that the tooling is in need of replacement.

