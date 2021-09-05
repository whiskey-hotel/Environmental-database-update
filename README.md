# Database update

## Background

The database at my current place of work is comprised of thousands of client and project data. It is used by several departments on a daily basis and has proven to be crucial in maintaining efficiency during the workday.

The database is stationed in Microsoft access and up until recently, everyone in the company could access and manage it at will. This is obviously a major security issue, so the DBA limited access to a handful of people to update on a regular basis.

While this is great for keeping the company safe, this severely limited my department’s ability to keep our project information up-to-date. Particularly because the database managers did not consistently update our worksheets. This meant that the hundreds of reports that my team wrote had no support from the database, therefore, productivity was hurt for several months.

After trying to explain the issue to our DBA and IT department, I decided to find a solution on my own.

## Solution

Since our database worksheets are copied into excel worksheets, my first attempt was to use VBA. The script worked ok, but I found it to be kind of clunky. I had a bit of experience with VBA in the past, but not enough to make something fast.

After fumbling through google for some of the quirks of VBA, I happened upon the Python library called **openpyxl**.

I have been using Python a lot recently, so I was pretty happy to learn that I could manage excel worksheets with Python.

## How it works

The basic idea of this script is to pull information from other active data worksheets and replace or add that information into my department’s worksheets. In this way, whenever an update is made by our DBA, I can simply run this script and update all of our data.

To avoid using up too much of the computer’s memory from read and write functions, all of the pertinent data is read into an array at the beginning of execution. The information is compared, altered and then appended to the end of the worksheet. While the read function is executed once. The write function occurs at the end of each iteration.

The script then looks in our database for projects that do not have contract numbers. If a project doesn’t have a contract number, it makes it really hard for my team to work with and identify the correct project information. So, the script compares a number of columns with each other (five (5) columns as of 9/5/21). If the information in our worksheet makes a match with all of the identified columns of the active worksheet, the contract number of the active worksheet is copied into our worksheet.

The script then checks for new projects that have not been added to our database by comparing the projects in our worksheet to those of another active worksheet. I did this by comparing contract numbers (since this is always unique). If a contract number is not found, then the contract number and all of its associated information is then added to our worksheet.

The information is then saved and logged for reference.

## Pitfalls

One pitfall is the way that _missing projects_ are identified. It works great for projects that have contract numbers. However, there are some projects that don’t have a project number yet, or are using a temporary contract number. If a contract number is missing, then those projects fly under the radar. If a temporary contract number is used, the script will add it to our worksheet, but when a legitimate number is added, we will effectively have duplicate projects in the database.

Another pitfall is the way in which _missing contract numbers_ are identified. If all of the client information is present in both worksheets, this process works out (this is true in most cases). But if there is missing client data in either worksheet, then the contract numbers will go unfilled for our worksheet.

## Conclusion

While the script is not perfect, it has brought the productivity of my team back to where it was prior to the change in database protocol.

## Upcoming updates

There are a few features that I need to add which may require a complete overhaul of the program logic:

-   Adding a check for empty client info in our worksheet.
-   Adding a check for updated client info in our worksheet.
-   Adding a check for missing contract numbers in the active worksheets.
