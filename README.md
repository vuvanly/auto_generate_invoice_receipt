# Background

Before, all employees have to create their own invoice and receipt file when they receive the salary. 
So I used this tools to help the company generate and print automatically so the employees don't have to do such boring task any more

# Idea

I used phpdocx and phpExcel library to create docx and excel file from the template.
The salary from employee will be loaded from CSV file and generate invoice and receipt file.

After that, the program will call the command of WPSOffice to connect to printer and print the documents.
