# EasyPayRoll

A simple Powershell script that allows solopreneurs to run payroll in minutes instead of hours.

1. Open the script and update your email address: $defaultemail = "[youremail@email.com]"

2. Run the script.  It will set up the necessary folders.

3. Access the [google docs spreadsheet template](https://docs.google.com/spreadsheets/d/176l4xfrCLiFInZnm5kYxbe6vaMFwkLZjlRWWKZwTW5M/edit?usp=sharing).

4. Make your own copy of the google docs spreadsheet and save it.

5. In your saved google docs spreadsheet, enter the contractor information (one row per contractor).  
	- Name: Contractor first and last name.
	- Email: Contractor email.
	- Pay Period Start: The starting date for the pay period.
	- Pay Period End: The ending date for the pay period.
	- Quantity: How many "units" of the specific task the contractor completed.
	- Rate: Contractor pay rate.
	- Extra: This column is for bonuses.  
	- Total Amount: This is automatically calculated: (Quantity * Rate) + Extra, and is the amount the contractor will be paid.
	- Notes: Put notes about performance, specific tasks, etc.

6. Once you've entered the information, save it as a .CSV file in your default download folder.

7. Run the script again.  It will ask you to confirm the information. Press Y to send the emails, press N kill the script and start over.


# How to use the script regularly

1. Update your spreadsheet with the relevant information.
    
2. Run the script.
