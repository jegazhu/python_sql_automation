import pyodbc
import pandas as pd
import os
from datetime import datetime
from plyer import notification

conn = pyodbc.connect(
    "Driver={SQL Server Native Client 11.0};" #feel free to check your driver with 'select @@SERVERNAME'
    "Server='your server connection goes here';"
    "Database='your database goes here';"
    "Trusted_Connection=yes;"
    )

#to insert data into table
#cursor=conn.cursor()
#cursor.execute("insert into Persons(Email, First_Name, Last_Name, Birth_Date, Address, State, City, Zip_Code) values(?,?,?,?,?,?,?,?)",('NA','Joe','Doe','2008-10-24','Malvin', 'Montevideo', 'Montevideo', '11400'))
#cursor.commit();

#to read data from the table
#cursor=conn.cursor()
#cursor.execute("select FS.[Sale Key], FS.[City Key], FS.[Stock Item Key], DC.City from Fact.Sale as FS inner join Dimension.City DC on FS.[City Key] = DC.[City Key] where DC.City = 'Lompoc' ")
#rows = cursor.fetchall()
#for row in rows:
#    print(row, end='\n')
sqlQuery = "select FS.[Sale Key], FS.[City Key], FS.[Stock Item Key], DC.City from Fact.Sale as FS inner join Dimension.City DC on FS.[City Key] = DC.[City Key] where DC.City = 'Munich' "

#
df = pd.read_sql(sql = sqlQuery, con = conn)

#
df.to_csv(os.environ["userprofile"] + "\\OneDrive\\Escritorio\CopyData\\" + "SQL_DataXtrac_" + datetime.now().strftime("%d-%b-%Y %H%M%S") + ".csv", index = False)

#
notification.notify(title = "Report Status",
                    message = f"Sales data has been successfully saved into Excel.\
                    \nTotal Rows: {df.shape[0]}\nTotal Columns: {df.shape[1]}",
                    timeout = 10)
