
<img src="assets\inventory banner.png" alt="Inventory Management Banner" />


<div align="center">

# **Inventory Management System**


</div>

<div align="justify">

 This project is a focus on the fulfilments of a typical warehouse/inventory anaysis and managemnet for a client-contractor relationship. Interfaces are provided for the users to manage transactions occuring in the warehouse and obtain reports of the inventory data. 
 Use this readme file to understand the project and the methods employed. If you would like to collaborate on data-related project, you can contact me on israelbssy@gmail.com.


Please, feel free to contribute to the project any way you can. Cheers!

 
</div>


<div align="center">

![Python version](https://img.shields.io/badge/microsoft%20power%20BI-darkgrey?style=for-the-badge)
![Python version](https://img.shields.io/badge/-microsoft%20SQL%20SERVER-yellow?style=for-the-badge)
![Excel Version](https://img.shields.io/badge/EXCEL%20VERSION%20-365-blue?style=for-the-badge)
![GitHub last commit](https://img.shields.io/github/last-commit/BasseyIsrael/inventory-management-system?style=for-the-badge)
![Type of ML](https://img.shields.io/badge/%20-analysis%20and%20management-red?style=for-the-badge)
![License](https://img.shields.io/github/license/BasseyIsrael/INventory-management-system?style=for-the-badge)

[![Open Source Love svg1](https://badges.frapsoft.com/os/v1/open-source.svg?v=103)](https://github.com/ellerbrock/open-source-badges/)

For the Badges [source](https://shields.io/)



</div>

## **Author**

- [Israel Bassey](https://github.com/BasseyIsrael)

## **Table of Contents**

  - [Introduction](#introduction)
  - [Data Source](#data-source)
  - [System Requirements](#system-requirements)
  - [Solutions Employed](#solutions-employed)
  - [Limitation and Improvement Opportunities](#limitations-and-improvement-opportunities)
  - [Explore Dashboards](#explore-dashboards-view)
  - [Conclusion](#explore-dashboard)
  - [License](#license)

<div align="Justify">


# **Introduction**

In the context of management involving a contractor and clients assets, the importance of properly managing the activities and state of the inventory items under your care cannot be over-emphasized. It is necessary to ensure that the effects of the addition and removal of items in the warehouse be properly captured and possibly used for further analysis. 
In some cases, the processes involved in ensuring the meeting the need can be challenging especially without teh use of a third-party installed software. It is for this purpose that this system was developed for an in-house continous management of the warehouse inventory.

For this project, various variables are at play, but they all lie within the sphere of:

- what happens with an item being added, 
- what happens when it is being removed, and 
- what happens when nothing is going on?

The key objective of this project is to build an interconnected inventory handling (for the contractor) and then a reporting system (for the client).

# **Problem Definition**

The problems leading to the need for the development of an end-to-end solution lie in the use of the existing system.

The current system is a manually operated system in which the users (contractors) manually change numerical values (by adding or subtracting) on a spreadhsheet following any activity in the inventory. This process requires periodically counting the quantity of items available (for each 169 unique items) to reconcile what is contained in the workbook. This process comes with disadvantages as listed below:

- Process is time consuming as for every activity, the master list needs to be manually scanned and a number needs to be added or subtracted to the existing number by the user depending on the activity.
- Lack of a actual database which makes it tedious to manage data growth especially the logs created from each activity.
- More hours consumed to generate required reports.
- Lack of logs to keep track of the activities being carried out in the system.
- Lack of accountability as there are no evidences of things happening at the warehouse.

Alongside these disadvantages, there is also the problem of inability to install third party applications on a dedicated system.
All of the solutions implemented in this project are used to tackle these problems.

The newly developed system will allow the users to:

- Seemlessly add or subtract items on the list using just selection and button clicks (This is the only activity needed by the user).
- Query the list by searching for specific items, setting thresholds, and easily finding levels.
- keep a log of the additions or removal of items in the inventory.
- Maintain notification summary of every single activity going on in the system.
- Upload and store (manually and automatically) all new data to a database which can be accessed by any BI tool.
- Send weekly emails of reoprts on the current state of the items in the inventory.
-  Save dated master lists for easy auditing.
- Understand the activity trend in the warehouse based on daily timing.
- Readily know all the items that fall below a threshold or restock value.
- Obtain automatic data updates on the BI tool using the Direct Query option in MS Power BI.
- Easily audit the destination of an item when it is removed from the inventory. 



# **Data Source**

This was a real business solution provided, hence actual inventory data was used to develop the system. The data in this repository is however anonymized for specific reasons.


# **System Requirements**

The system developed in this projet make use of the following platforms:

- Microsoft Excel - For the development of the data input part of the project. This also carries a lot of the automation codes in VBA used in the whole process.

- SQL Server - For the storage of the data generated in all the processes and for loading the generated data into the Business Intelligence tool.
- Microsoft Power BI - The most important part of the management system is the part that provides the report on the inventory to the Clients. The data fed into the system is gotten from the database and connected as live data to Power BI using Direct Query. It is for this reason that SQL Server was used. 
- Power Automate - In an earlier version of the system, Power Automate was used to trigger the data upload to the database and the dashboard refresh. This was however switched to a VBA script that does the data upload automatically upon every workbook save instance.

# **Solutions Employed**

The solutions employed start with the MS Excel workbook which can be found [here](https://github.com/BasseyIsrael/Inventory-Management-System/tree/main/Dashboards). The excel workbook formulas will not be explained in this readme however, the VBA codes will be highlighted. To handle the first problem of manually scanning the master-list, Excel Dynamic Array functions were used. The dilemma was to choose between using Dynamic Array function and using Excel VBA. Dynamic functions was what I settled with for easy data reference in the future.

To further the process, following every item added, a script is run to add a data log on the item added, when it was added and the quantity. This partly solves the problem of accountability. The script can be seen below.

```bash

add_answer = MsgBox("Hello " & Excel.Application.UserName & ", You are about to add a new item to the inventory. Please confirm that this addition is correct before proceeding.", vbYesNo + vbQuestion + vbDefaultButton2, "Add Item")
If add_answer = vbYes Then

    noi = Sheets("Dashboard").Range("J22").Value
    noa = Sheets("Dashboard").Range("N22").Value
    dat = Sheets("Dashboard").Range("M19").Value
    
            
    iRow = Sheets("Items Added").Range("B1048576").End(xlUp).Row + 1
    
    
        With ThisWorkbook.Sheets("Items Added")
        
            .Range("B" & iRow).Value = iRow - 6
            .Range("C" & iRow).Value = noi
            .Range("D" & iRow).Value = noa
            .Range("E" & iRow).Value = dat
            '.Range("F" & iRow).Value = 5
            
        End With
```

Following every added item, a notification is created that stated the kind of activity that has been carried out. The notification is done for both removal and addition.

```bash

name_of_item = Sheets("Data Summary").Range("AJ4").Value
no_of_items = Sheets("Data Summary").Range("AK4").Value
            
add_message = name_of_item & " (" & no_of_items & ") was added to the inventory on " & dat & "."
        
notif_row = Sheets("Notifications").Range("B1048576").End(xlUp).Row + 1


    With ThisWorkbook.Sheets("Notifications")
    
        .Range("B" & notif_row).Value = notif_row - 6
        .Range("C" & notif_row).Value = dat
        .Range("D" & notif_row).Value = add_message
    
            End With
```

What happens with the item addition also happens when items are removed from the warehouse. Because the items removed from the warehouse are usually carried to a work location (field) by a work barge, these information are also captured to aid full accountability alongside the notification. The use of the notification is to give the user the ability to send mail updates to relevant parties to keep them up to date with the activities. The script to add an item to the log list and to also create a notification for the activity are seen below:

```bash 
Dim remove_answer As VbMsgBoxResult

remove_answer = MsgBox("Hello " & Excel.Application.UserName & ", You are about to remove an item from the inventory. Please confirm that this action is correct before proceeding.", vbYesNo + vbQuestion + vbDefaultButton2, "Add Item")
If remove_answer = vbYes Then

tnoi = Sheets("Dashboard").Range("J35:M36").Value
tnoa = Sheets("Dashboard").Range("N35:O36").Value
tdat = Sheets("Dashboard").Range("M32").Value
tloc = Sheets("Dashboard").Range("J38:L39").Value
tbar = Sheets("Dashboard").Range("M38:O39").Value
tper = Sheets("Dashboard").Range("J41:O42").Value

tRow = Sheets("Items Removed").Range("B1048576").End(xlUp).Row + 1


    With ThisWorkbook.Sheets("Items Removed")
    
        .Range("B" & tRow).Value = tRow - 6
        .Range("C" & tRow).Value = tnoi
        .Range("D" & tRow).Value = tnoa
        .Range("E" & tRow).Value = tbar
        .Range("F" & tRow).Value = tper
        .Range("G" & tRow).Value = tloc
        .Range("H" & tRow).Value = tdat
            
        End With
```

```bash

name_of_ritem = Sheets("Data Summary").Range("AJ5").Value
no_of_ritems = Sheets("Data Summary").Range("AK5").Value
name_of_rloc = Sheets("Data Summary").Range("AL5").Value
name_of_rbar = Sheets("Data Summary").Range("AM5").Value
            
add_rmessage = name_of_ritem & " (" & no_of_ritems & ") was removed from the inventory to " & name_of_rloc & " by " & name_of_rbar & " barge on " & tdat & "."
        
notif_row = Sheets("Notifications").Range("B1048576").End(xlUp).Row + 1


    With ThisWorkbook.Sheets("Notifications")
    
        .Range("B" & notif_row).Value = notif_row - 6
        .Range("C" & notif_row).Value = tdat
        .Range("D" & notif_row).Value = add_rmessage
    
    End With

```
You would notice that these scripts are not wrapped in a procedural call. This is because they are all part of a larger procedure.

To generate a quick report, a "Send Email" option is provided. The script in my productivity toolpack on Excel [here](https://github.com/BasseyIsrael/Excel-VBA-Scripts/blob/main/VBA%20Scripts/Save%20and%20send%20email.vb) was used. This script allows a user to send an email with a pdf attachment of one or more of the worksheets. This simple task saves the operator the hassle of having to manually save said attachments and go ahead to send the email to the required party, as this is done in one click (Maybe 2).

The next issue to be handled was the updating of data to the BI tool for reporting and further analysis. Three methods were proposed to handle this:

- Use the scheduled refresh feature for a sharepoint workbook.
- Add a Power Automate flow to trigger the data refresh.
- Write a VBA script to upload the data to a database following a save instance on teh workbook and connect the database to the BI tool using Direct Query.

The first method involves the following:
- Upload the workbook to onedrive 
- Obtain the onedrive path 
- Paste it on Power BI as a web link. Power BI detects what is being done and sees the excel workbook you are trying to upload. 
- Verify your credentials
- Load your workbook
- Publish to Power BI Service
- Schedule your data refresh as uch as you want. 
- Load your report to web.

The second method was not ideal for this project as the dashboad is a shared dashboard and the source file needs to be on the machine performing the refresh.

The third method is my personal favourite as it proves to be stable for the most part. It involved creating a database connection in Excel (database needs to be online and hosted), upload most recent data to the database, close connection, and let Direct Query do the rest. 
For this project, only 7 tables were created, so it was fairly easy to write the needed script for each one of them. The code can be seen below. It should be noted that this procedure works as trigger, so it should be saved in the "This Workbook" part of the VBE.

```bash
Private Sub Workbook_AfterSave()

End Sub

    Dim conn As New ADODB.Connection
    Dim TableNAme As String
    Dim sqlstr As String
    Dim re As ADODB.Recordset
    
    Dim CRow As Long, TRow As Long
    
    Set conn = New ADODB.Connection
    
    conn.Open "DRIVER={SQL Server}" & ";SERVER=" & "Your Server Name" _
    & ";DATABASE=" & "Your Database Name" _
    & ";UID=" & "Your User ID" _
    & ";PWD=" & "Your Password"
    
    Set rs = New ADODB.Recordset
    
    TRow = Sheets("Sheet Containing Table").Range("A" & Rows.Count).End(xlUp).Row
    
    For CRow = 1 To TRow
        
        Col_1 = Sheets("Sheet Containing Table").Range("A" & CRow).Value
        col_2 = Sheets("Sheet Containing Table").Range("B" & CRow).Value
        col_3 = Sheets("Sheet Containing Table").Range("C" & CRow).Value
        col_4 = Sheets("Sheet Containing Table").Range("D" & CRow).Value
        col_5 = Sheets("Sheet Containing Table").Range("E" & CRow).Value
        col_6 = Sheets("Sheet Containing Table").Range("F" & CRow).Value
        
        sqlstr = "INSERT INTO" & "Table_name" & "VALUES('" & Col_1 & "'," & "'" & col_2 & "'," & "'" & col_3 & "......" ')" 'Ensure you format this line to how you want it to be used or just input an SQL command you want to execute here.
        rs.Open sqlstr, conn, adOpenStatic
        
    Next
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    MsgBox ("Upload Complete")

```

The BI dashboard carries the information that the client needs to see. The data summary, threshold analysis, search query, relaivity analysis are all present on the dashboard.

Since the data obtained from the Excel workbook is relatively fine as is, the direct query just involves calling each of the tables from the database with a few lines of script.

The BI dashboard was shared as an embeded webapp to reduce the number of "cooks" on the report.

# **Limitations and Improvement Opportunities**

- What is being used for data input and linked to the database is a protected excel workbook. Though this might be easy to use an ideal considering the situation, it may not be sustainable. For future work, it is advisable to use a stand alone software to handle the data input and inventory checks before moving to the BI platform.

# **Explore Dashboards** [(View)](https://github.com/BasseyIsrael/Inventory-Management-System/tree/main/Dashboards)

The Main Reporting Dashboard

<img src="assets\main dashboard.PNG" alt="analysis" />

The Dashboard showing activity logs

<img src="assets\data log.png" alt="analysis" />

Data Input Platform

<img src="assets\input interface.PNG" alt="analysis" />



# **License**

MIT License

Copyright (c) 2022 Israel Bassey

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Learn more about [MIT](https://choosealicense.com/licenses/mit/) license
