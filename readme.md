Simple app for Data entry For Mining Company Resources. 
A list of 'available' companies with website links is provided.
Also utilizes 'periodic table' for mineral selection and global countires slider for mine localtion selection.
Exports saved data to spreadsheet. Saves the data locally where ever the app resides. 

Option to open the target company website so you can search for their original data source. 
Input files

    Company ID & Websites: /Scot_data_beta.xlsx
    Elements Sorted & Prioritised: /Periodic_table.xlsx
    Country slider mapping: pycountry.countries
    
Output
    SCOT_RD.xlsx
    
Extra Resources: Function'save_data' uses py library "from pox.shutils import find " to save in your desktop CWD.


Note: If you run this in Jupyter notebook you might get a kernel crash when you exit the app. 
      This does not accur in VSCode.

GUI: 

<img width="897" alt="PYQT Data entry app screenshot 2022-03-24 at 21 25 07" src="https://github.com/user-attachments/assets/a5eecd61-58d8-469d-b070-74c7adea21ec">
