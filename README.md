# Waste scan graphs

Add measurements in [the master workbook](data/waste_scans_master.xlsx), then run one command to make Dutch graphs. You do not need to edit any code. Each event gets a folder in `results`; `results/combined` uses all events together.

## Set up a computer for the first time

1. Install [Python for Windows](https://www.python.org/downloads/windows/). During installation, allow the Python launcher (`py`) if asked.
2. To edit the workbook, use Microsoft Excel if you have it. Keep the workbook in `.xlsx` format when saving.
3. Copy this entire project folder to your computer. Do not move the workbook out of its `data` folder.
4. Open the project folder in File Explorer. Click the address bar, type `powershell`, and press **Enter**. A command window opens in the right folder.
5. Type these commands, pressing **Enter** after each one:

   ```powershell
   py -3 --version
   py -3 -m pip install -r requirements.txt
   py -3 src\waste_analysis.py
   ```

   The last command creates the graphs. Next time, you only need to double-click `make_charts.bat`, or run the last command again.

## Add or edit a scan event

1. Open `data/waste_scans_master.xlsx`. In **Events**, add a unique name and start/end dates, for example `ASEB Feb 10-13 2026`. 
2. In **Scans**, add one row for each weighed bag. Select its event name and enter institution, campus, building, environment, and collected stream. Use the header filters to review rows.
3. Enter the bag's **total weight and six material weights in grams**. The six materials should add up to the total. Bag size and fill percentage are needed for the volume graph. Notes and source fields are optional.
4. Save and close the workbook. Double-click `make_charts.bat`. Open the matching event folder in `results`. Run it again after corrections; old graphs in each event folder are replaced.

If you see an error with a row number, correct that row in the workbook, save it, and run the batch file again.

## Graph order and availability

For every event, the tool checks which streams have measured data. It makes volume and weight by location and a contamination-by-material graph, then only the relevant stream sections. Each available stream gets total-versus-composition bars by location for every measured environment, plus highest/lowest purity pies where two locations can be compared. The final scatter graph for each stream shows sample weight against purity. Residual also gets paper-in-Rest and paper-separate views, plus GFT contamination in kilograms and percentages.

February has GFT, paper, PD, and residual sections. November has only residual. The workbook has no tissue stream or tissue material measurement, so the requested tissue graph is skipped. It can be added when those measurements are recorded.

The graphs show **sampled bags**, not all waste produced. Volume uses bag fill × nominal capacity (S 30 L, M 70 L, L/L+ 140 L, XL 350 L). A bag without fill appears in weight graphs but not volume graphs. The **Summary inputs** sheet holds measured subtotals when individual bag rows are unavailable; these are included in the graphs.
