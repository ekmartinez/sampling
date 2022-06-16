import pandas as pd
import numpy as os

#Risk table
risk_factor = 3
df = pd.DataFrame([[0, 100000, 0, .04], 
                    [100000, 500000, 2500, .02],  
                    [500000, 1000000, 8500, .01],
                    [1000000, 2999999,10000, .009]],
                    columns=['Low', 'High', 'Base', 'Percentage'])

#Creates Data Frame from clipboard data,
# just make sure your numeric column is named "Amount"                   
grab = pd.read_clipboard()
exp = grab['Amount'].sum()

def selector(level):
    """Selects random sample based on materiality"""
    output = f"{os.path.abspath(__file__)}\\output_file.xlsx"
    materiality = df.at[level, 'Base'] + df.at[level, 'Percentage'] * exp
    sample = (exp / materiality) * risk_factor
    sample_selection = grab.sample(n=int(sample), replace=True)
    sample_selection.to_excel(output)

#If total of numeric column is between level.
if exp > df.at[0,'Low'] and exp < df.at[0,'High']:
    selector(0)
    
elif exp > df.at[1,'Low'] and exp < df.at[1,'High']:
    selector(1)
    
elif exp > df.at[2,'Low'] and exp < df.at[2,'High']:
    selector(2)
   
elif exp > df.at[3,'Low'] and exp < df.at[3,'High']:
    selector(3)
    