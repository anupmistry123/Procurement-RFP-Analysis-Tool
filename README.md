# procurement_analysis_tool
Procurement Analysis Tool Code Library
########################################
- This is a collection of some of my SQL and VBA code. The purpose is to showcase my formatting, coding style and the nature of the work I was doing.
- In line with good coding etiquitte, I tried to standardadise the formatting and add as much commentary as possible.
- Most of this code refers to an analytics tool that had an Excel / VBA frontend and a MySQL / MariaDB backend. The tool was to anaylyse procurement RFPs.
  - RFPs are "Requests For Proposals" and are standard practices in procurement when a company wants to find new suppliers ( they are also known as sourcing tenders).   - They are exercises carried out where a company invites suppliers of certain items /services to bid on a select list of required items from the company.
  - The company then receives a price list from each supplier, with the supplier filling out prices for the items / services they can price.
  - The company would then group all the suppliers together and then carry out analysis e.g. cheapest supplier for all items, standard prices across all suppliers, areas of the RFP that lack sufficient coverage.
- The user would input "raw" data and "analysis parameters" into the Excel. The user would then click buttons which upload the data into a MariaDB database. 
- Automated analysis would then be carried out and the results (both high level and granular) would be presented back to the user in the Excel file.
########################################
