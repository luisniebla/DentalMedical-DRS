# Import to SQL

Need to structure our data in a particular way to make it easier 
to access and manage.
0.) Import the Master
1.) Add a patient_valid_from column
2.) Add a patient_valid until column
3.) Add a patient_status column
    "Office Inactive"
    "DRS Inactive"
    "Above The Line Inactive"
    "Below the Line"
    "In Month Tab"
4.) Change the status column of master data
5.) Import month data, change the status to "In Month Tab" 

Now we have one big database in SQL for this campaign. We are able to 
track how it evolves, and how it becomes who it is.

When adding new data, I think we should do it this way:
1.) Move the old data to an archive table, set its valid_until date to today
2.) Move in the new data, with valid_from date to today

When changing data
1.) Move old data to archive table, set valid_until date to today
2.) Copy old data, change any necessary stuff, set valid_From date to today.


Export to CSV: DONE
Import CSV To SQL: In progress