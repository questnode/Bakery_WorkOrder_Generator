This project is done for a wholesale bakery.  The bakery offers 40+ types of bread prepared fresh on a daily bases.  Customers place their order by 12pm.  After the cut-off time, the operation staff must compile all orders and calculate the appropriate recipes for the production staff.  Production staff will then organize and arrange tasks amongst the team.

There are a few challenges with this workflow.
1. Due to the amount of products offered, it takes the staff a while to generate all recipes even when there are excel formula already set up.
2. Many products share same or similar components in their recipes.  These common components need to be parted out for batch production.  Therefore increasing the complexity of recipe generations.
3. Many products require precursors that need to be prepared 1 to 2 days in advance.  These precursors may also share same or similar components with other recipes.
4. Production staff takes time to sort and organize the tasks once recipes are generated.  

The objective of this project is as follows:
1. Reduce time and effort in generating all recipes
2. The ability to incorporate estimated precursors into daily recipes
3. Be able to divide each recipe into tasks by stages.  Eg. Mixing, Portioning, Forming, Baking.
4. Generate a daily work order package for each production staff.  The package includes all tasks and timing requirements, and recipes required on that day for him/her to complete production.

To accomplish these objective, the following steps are taken
1. Prepare Data
1.1 I set a folder to place the daily sales order from the company's ERP.  The daily sales order contains all products required for the next day
1.2 All recipes are compiled into a table
1.3 Production workflow is recorded and observed for 4 weeks.  The record is turn into two files
	1.3.1. Staff skillsets & scheduling
	1.3.2. Task stages & timing
1.4 I created a new Excel workbook "Work Orders.xlsm" to be used for generating the objectives.  It is first set up to do a few things:
	1.4.1. Load in daily sales order
	1.4.2. Allow input of precursors desired for that day
	1.4.3. Load in staff skillset & scheduling
	1.4.4. Load in task stages & timing
	1.4.5. Load in templates for individual staff work orders

2. Transformation using VBA in Excel
2.1 In the workbook, the daily sales order is scanned and parsed for all product requirements.
2.2 All quantities of products, sub-products, and precursors are derived and consolidated.  
2.3 From 2.2 all recipes needed for that shift are calculated.  These recipes are defined as task items.
2.4 All product are broken up into production stages.  Each product-stage is define as a task item.
2.5 All task items are allocated to appropriate production staff according to staffing schedule and the staff skillset.
2.6 Some items require overnight resting (storage) before they can be complete.  These are separated into another excel output to be incorporated the next day.
2.7 For items finished resting from the day before, the workbook imports them back for staff assignment and completion.
2.8 Each task items is assigned a time for completion based on 1.3.2.  
2.9 All task items are group by staff assigned, and sorted by timing.
2.10 All task items are loaded into work order template in 1.4.5 and saved to a new folder.
2.11 The generated work order package is sent to a printer.
