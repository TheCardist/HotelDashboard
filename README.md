# Hotel Configuration Dashboard
## Problem
A group of stakeholders voiced the need for hotel settings to be in a single view to make it easier to review how their portfolio of hotels is set up. They asked for a data dump file with all the data but this would be hundreds of thousands of unreadable data that would need to be formatted for use.

## Solution
The dashboard I created does have the data dump of information but to make it usable the main tab in Excel uses filter, xlookups, and various other formulas to create a dashboard-like view for easier use. The user only needs to type in the hotel name and it will populate the page with all the current settings.

Python was implemented into this project to query the database, add the query results to the Excel template, create a new version of the document and email it out to the stakeholders.
