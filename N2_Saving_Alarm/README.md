## Nitrogen Gas Saving Alarm
In TI's factory, there is a solder reflow machine. The solder reflow machine would take in nitrogen gas to help push oxygen gas away, preventing the solder from oxidizing. To do so, a machine (Machine A) would send solder paste dies to another machine (Machine B), and nitrogen gas would then be pumped into Machine B. Normally, when operating, both machines would be in Running state. During downtimes, both machines should be in Waiting state. However, the states of Machine B need to be switched manually by factory workers while Machine A can be switched automatically by the machine itself. There were multiple occasions where factory workers forgot to set Machine B from Running state to Waiting state when Machine A is in Waiting state, thus pumping nitrogen gas while no solder paste dies are being sent into Machine B, causing waste of nitrogen gas and money. While there are sensors detecting the states of both machines and would output the states onto an Excel file every 4 minutes, there isn't an automatic system that detects waste of nitrogen gas, i.e Machine A in Waiting state and Machine B in Running state. Therefore, the goal of this project is to design a program that helps detect nitrogen gas waste and also send alert emails to engineer managers if there's waste happening.

The Python script here reads the Excel file, detect the color codes of the cells (yellow cells represent Waiting state, green cells represent Running state), and determine whether there is waster happening or not. If it detects that nitrogen gas is being wasted, it would send an alert email first to the engineer manager. If the waste is still happening after another chosen timeframe, a second alert email would be sent. Users can also set downtimes for emails so that the script won't spam emails to the receivers. Users set downtime, desired timeframe used for determining whether the waste is happening in an Excel file. Users can also customize recipients, subjects, and bodies of emails on the same Excel file.

To run the script periodically (every 5 minutes), I use the Windows Task Scheduler to set it to run the script every 5 minutes.

Below is the logic flow of the script:
<br/>
<br/>
<img src="pics/LogicFlow.PNG" width="75%">
<br/>
<br/>
