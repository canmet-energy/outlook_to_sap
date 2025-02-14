# Outlook Data Anaylsis / SAP Timesheet helper

## Why
* Do you hate entering your timesheets in SAP? 
* Do you find it is a one-way flow of information that you can't really use to manage your own time?
* Do you struggle to figure out where your time has been spent?
* Do you prefer **Outlook** to **SAP** to record your time?
* For each quarter / Performance review would you like to know where your time was spent?


You can use this python script to:

    1. scan your Outlook Calendar
    2. compare appointments that you have tagged to your project list
    3. add current week to your clipboard
    4. paste your week into SAP. 

## Prequisites
* Install [Python](https://www.python.org/ftp/python/3.7.4/python-3.7.4-amd64.exe) for windows 3.74.  
* If you are wanting to learn more about Python, I would recommend installing [PyCharms Community Edition](https://download.jetbrains.com/python/pycharm-community-2019.2.exe)

## Simple Way


1. Before we start, clone this repository directly to your C: from git. You must have Git installed.
```bash
cd c:/
git clone https://github.com/canmet-energy/outlook_to_sap.git
```

2. Edit your projects list by modifying the Excel projects.xlsx in the resources folder. The default is set to mine like this. Note the short project_nicknames that I gave for each project that we will use later.
![alt text](https://github.com/canmet-energy/outlook_to_sap/raw/master/images/Excel.png)


3. Use Prefixes for your events/meetings in outlook to match the nicknames that you have defined in your projects.xlsx file and optionally your task number. The syntax is <project_nickname>:<task_digits>:<Subject> For example to charge to the **Pathyways Project**. I prefix any timeslot with a **PW:** prefix I defined in my Excel file. If I had a btap task number.. it would be **PW:<task_digits>:subject**

Note: BTAP Task users.. you can enter your btap_task number as well. It the prefix it with 
![alt text](https://github.com/canmet-energy/outlook_to_sap/raw/master/images/outlook.png)

3. Open up SAP and go to the week you wish to examine. Get the weeknumber.
![alt text](https://github.com/canmet-energy/outlook_to_sap/raw/master/images/sap_week_number.png)

4. Edit the outlook_to_sap\sap_copy_week.bat and change the weeknumber and year.
![alt text](https://github.com/canmet-energy/outlook_to_sap/raw/master/images/bat_file.png)

5. Run outlook_to_sap\sap_copy_week.bat. This will copy the week into your clipboard.
![alt text](https://github.com/canmet-energy/outlook_to_sap/raw/master/images/command.png)

6. Go to your SAP week Data Entry area, select the first Cost Center field and hit CTRL+v.
![alt text](https://github.com/canmet-energy/outlook_to_sap/raw/master/images/sap_paste1.png)

7. You now have entered the week. Review the hours so it makes sense.  
![alt text](https://github.com/canmet-energy/outlook_to_sap/raw/master/images/finish.png)

Note: SAP supports Python scripting. I've requested that Agriculture Canada turn that feature on. That way you will not need to even paste into your timesheets in the near future. 



