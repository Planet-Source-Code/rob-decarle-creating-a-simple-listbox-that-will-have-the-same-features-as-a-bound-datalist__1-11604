<div align="center">

## Creating a simple listbox that will have the same features as a bound datalist box


</div>

### Description

The purpose of this example is to connect/display database records (using ADO) with a simple listbox. This is a work around to using bound controls. You will find your apps will run faster.
 
### More Info
 
You will need to draw a listbox on a form. You will need to create an Access database and an ODBC DataSource.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rob deCarle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rob-decarle.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rob-decarle-creating-a-simple-listbox-that-will-have-the-same-features-as-a-bound-datalist__1-11604/archive/master.zip)





### Source Code

```
'**********************************************
'*Put the following code in any function.
'*This code opens database connection/recordset.
'*I have connected via datasource
'**********************************************
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
cn.ConnectionString = "Provider=MSDASQL.1; _
Persist Security Info=False; _
Data Source=DSG_Input"
cn.Open
'*****************************************
'*SQL Statement to extract all customers*
'*from the database
'*****************************************
sql = "Select First_Name, Cust_ID from Customer _ Order by First_Name"
Set rs = cn.Execute(sql)
'*****************************************
'**Populates the listbox**
'*****************************************
  With List1
    Do While Not rs.EOF
    .AddItem rs("First_Last")
    rs.MoveNext
    Loop
  End With
'**********************************************
'*You now have a listbox containing the records
'*from your database
'**********************************************
'**********************************************
'*You will create an array that is dynamic to
'*your recordset. This will keep track of
'*the primary key as a boundColumn would in a
'*datalist box. This is for the purpose
'*of relational databases.
'*You will create the array the same size as the
'*listIndex count (number of records in
'*listbox).
'**********************************************
rs.movefirst
ReDim array1(List1.ListCount) As String
'*********************************************
'*This will now populate the array which is a
'*mirror image as the listbox, but with the
'*primary key.
'*********************************************
For i = 0 To List1.ListCount - 1
  array1(i) = rs("Cust_ID")
  rs.MoveNext
Next i
'**********************************************
'*We have now completed the listbox. You can
'*use this listbox the same way as you would a
'*datalist box. The following code will explain
'*how.
'***********************************************
'************************************************
'*To access the primary key relating to each
'*record in the list, put the following code in
'*the listbox "Click()" event. This explains how
'*to access the primary key stored in the array.
'************************************************
'**********************************************
'*list1.listIndex explains with record in the
'*list was clicked on. You use this to find
'*where in the array the primary key is stored.
'**********************************************
Private Sub List1_Click()
Dim Primary_1 as string
  Primary_1 = array1(list1.listIndex)
  Msgbox Primary_1
End Sub
'***********************************************
'*Conclusion*
'*Although this isn't as convenient as setting up
'*a bound datalist control, you will find it
'*will speed up things when using a large
'*database file.
'************************************************
```

