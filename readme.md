# Catalog Updater

A python script developed to aid in transferring the important information from a distrubuter's **Catalog** to a Seller's **Cost Price Files** and **Resale Price Files** Template.

-*Made to work with excel .xlsx files only*-

This was created to work for my own job but it's published for anyone to take a look at and use. I can't know if it will be useful for many people but I do believe this code can be edited fairy easily to do a lot of different CopyPasting tasks in Microsoft Excel with multiple columns and pages.

--- 

# How to Use:

1. Get your Template and put it's file path in main.py line 10
2. Put distributer catalogs inside **/Inbox** folder
3. Inside the function *dialog* there is a list called lst[] where you can enter the 5 column letters of the original price files you want to copy to the template file
4. Run main.py
5. Check **/Outbox** folder

---

# Todo

- ~~Before running, display imported files~~
- ~~Add Color to print statements for readability in Terminal~~
- Perhaps manually confirm which columns match up in Terminal 
- Automatically import Template File

---

Credit to *Openpyxl* for the fantastic Excel + Python intergration
