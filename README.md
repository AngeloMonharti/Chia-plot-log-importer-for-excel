# Chia-plot-log-importer-for-excel
Imports information from the default chia log location into excel, useful for small time farmers and noobs (Only works with English plot logs)

Open excel -> options -> customize ribbon -> make sure "Developer" is checked on the right-hand list -> OK -> developer (new tab appears in the top menu bar) -> Visual Basic (first button on the left under the menu bar) -> Tools -> click "references" in the drop down -> scroll down until you see "Microsoft Scripting Runtime" -> click the checkbox -> OK -> now go to Insert -> click "module" from the drop down -> copy and paste the code into the module window -> hit green arrow under the "debug" menu while the module is selected to run the macro. Alternatively, go back to the excel sheet and beside the Visual Basic button in the Developer tab that you clicked earlier, there is a "macros" button. Click that, select ChiaPlotInfo, click run. If you want to save the file, save as .xlsm (not .xlsx) so the macro module doesn't get deleted.



Make sure you edit your user profile name in the Plot directory or it won't be able to find the plot logs -> PlotDir = "C:\Users\USER\.chia\mainnet\plotter"

Look I literally have no coding knowledge I'm just an office email monkey who learned how to do VBA so critique the code all you want I don't care

Here are the columns in order:
Plot Name, Plot Size, Temp Directory Used, Final Directory Used, Bitfield Enabled (Y/N), Memory, Buckets, Threads, Stripe, Start time, Phase 1 Complete, Phase 2 Complete, Phase 3 Complete, Phase 4 Complete, Phase 1 Duration (minutes), Phase 2 Duration (minutes), Phase 3 Duration (minutes), Phase 4 Duration (minutes)
