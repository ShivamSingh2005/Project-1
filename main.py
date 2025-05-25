'''
In terminal execute:
    uvicorn main:app --reload --port 9090 
to run the API.

Then click on the link given at the end of README.md file to test the API on Postman.
'''

#Import Modules
from fastapi import FastAPI
from pydantic import BaseModel
import pandas as pd

#Read the file
df = pd.read_excel("./Data/capbudg.xls")

# Parse Data From capbudg.xls
tables = {}
tables['INITIAL INVESTMENT']=df.iloc[2:9,0:3:2]
tables['CASHFLOW DETAILS']=df.iloc[2:6,4:7:2]
tables['DISCOUNT RATE']=df.iloc[2:10,8:11:2]
tables['WORKING CAPITAL']=df.iloc[11:14,0:3:2]
tables['INITIAL INVESTMENT Details']=df.iloc[23:30,0:2]
tables['Investment Measures']=df.iloc[52:55,1:3]
tables['GROWTH RATES']=df.iloc[17:19,0:12] 
tables['GROWTH RATES'] = tables['GROWTH RATES'].drop(df.columns[[1,2]], axis=1)
tables['SALVAGE VALUE']=df.iloc[32:34,0:12]
tables['SALVAGE VALUE'] = tables['SALVAGE VALUE'].drop(df.columns[[1]], axis=1)
tables['OPERATING CASHFLOWS']=df.iloc[36:50,0:12]
tables['OPERATING CASHFLOWS'] = tables['OPERATING CASHFLOWS'].drop(df.columns[[1]], axis=1)

app = FastAPI()

#Root 
@app.get("/")
async def root():

    """
    The function executes when server calls the root url.

    Returns
    -------
    dict
        About the Project
    """

    return  {"Name": "Shivam Kumar Singh",
              "Project Name": "FastAPI Excel Processor Assignment",
              "Version": "Without ChatGPT"
             }

#List_Tables    
@app.get("/list_tables")    
async def list_tables():
    """
    Lists all the tables in capbudg.xls

    Returns
    -------
    dict
        A list containing all the table names.

    Example
    -------
    Terminal: uvicorn main:app --reload --port 9090
    Go to: http://127.0.0.1:9090/list_tables
    Output:
    {
        "tables": [
            "INITIAL INVESTMENT",
            "CASHFLOW DETAILS",
            "DISCOUNT RATE",
            "WORKING CAPITAL",
            "INITIAL INVESTMENT Details",
            "Investment Measures",
            "GROWTH RATES",
            "SALVAGE VALUE",
            "OPERATING CASHFLOWS"
        ]
    }
    """
    return {"tables": list(tables.keys())}

#Get Table Details
@app.get("/get_table_details/{table_name}")
async def table_details(table_name : str):
    """
    Gives the table details containg its name and row names.

    Parameters
    ----------
    table_name : str
        The name of the table

    Returns
    -------
    dict
        The table name and the list containing row_names in the table.

    Raises
    ------
    Invalid Table Name
        If Table Name not found in capbudg.xls

    Examples
    --------
    Terminal: uvicorn main:app --reload --port 9090
    Go to: http://127.0.0.1:9090/list_tables
    Output:
    {
        "table_name": "DISCOUNT RATE",
        "row_names": [
            "Approach(1:Direct;2:CAPM)=",
            "1. Discount rate =",
            "2a. Beta",
            " b. Riskless rate=",
            " c. Market risk premium =",
            " d. Debt Ratio =",
            " e. Cost of Borrowing =",
            "Discount rate used="
        ]
    }
    """
    try:
        table_det = {}
        table_det['table_name']= table_name
        table_det['row_names']= list(tables[table_name].iloc[:,0])
        return table_det
    except:
        return {"Invalid Table Name":f"No table with name <{table_name}> found."}

#Get Row Sum
@app.get("/row_sum/{table_name}/{row_name}")
async def row_sum(table_name: str, row_name: str):
    """
    Calculates the sum of all the numbers row.

    Parameters
    ----------
    table_name: str
        The name of the table
    row_name: str
        The name of the row

    Returns
    -------
    dict
        The dictionary containing table name, row name and the sum of the values int that row.

    Raises
    ------
    Invalid Table Name
        If Table Name not found in capbudg.xls

    Invalid Row Name
        If Row name not found in the table but table namee is valid.

    Examples
    --------
    Terminal: uvicorn main:app --reload --port 9090
    Go to: http://127.0.0.1:9090/row_sum/INITIAL%20INVESTMENT/Initial%20Investment=
    Output:
    {
        "table_name": "INITIAL INVESTMENT",
        "row_name": "Initial Investment=",
        "sum": 50000
    }
    """

    row_sum = {}
    row_sum['table_name']=table_name
    row_sum['row_name']=row_name
    sum = 0
    index = 0
    try:
        for row in tables[table_name].iloc[:,0]:
            if (row == row_name):
                sum_row = tables[table_name].iloc[index,1:]
                break
            index+=1
        else:
            return {"Invalid Row Name":f"No row with name <{row_name}> found in table <{table_name}>."}
        for num in sum_row:
            sum+=num
        row_sum['sum']=sum
        return row_sum
    except:
        return {"Invalid Table Name":f"No table with name <{table_name}> found."}