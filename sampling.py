import pandas as pd
import numpy as np

#Risk based approach
risk_factor = 3

grab = pd.read_clipboard()

df = pd.DataFrame(np.array([[0, 100000, 0, .04], 
                            [100000, 500000, 2500, .02],  
                            [500000, 1000000, 8500, .01],
                            [1000000, 2999999,10000, .009]]),
                   columns=['Low', 'High', 'Base', 'Percentage'])

exp = grab['Amount'].sum()

if exp > df.at[0,'Low'] and exp < df.at[0,'High']:
    materiality = df.at[0, 'Base'] + df.at[0, 'Percentage'] * exp
    sample = (exp / materiality) * risk_factor
    sample_selection = grab.sample(n=int(sample), replace=True)
    sample_selection.to_excel('output_file.xlsx')
    
elif exp > df.at[1,'Low'] and exp < df.at[1,'High']:
    materiality = df.at[1, 'Base'] + df.at[1, 'Percentage'] * exp
    sample = (exp / materiality) * risk_factor
    sample_selection = grab.sample(n=int(sample), replace=True)
    sample_selection.to_excel('output_file.xlsx')

elif exp > df.at[2,'Low'] and exp < df.at[2,'High']:
    materiality = df.at[2, 'Base'] + df.at[2, 'Percentage'] * exp
    sample = (exp / materiality) * risk_factor
    sample_selection = grab.sample(n=int(sample), replace=True)
    sample_selection.to_excel('output_file.xlsx')
    
elif exp > df.at[3,'Low'] and exp < df.at[3,'High']:
    materiality = df.at[3, 'Base'] + df.at[3, 'Percentage'] * exp
    sample = (exp / materiality) * risk_factor
    sample_selection = grab.sample(n=int(sample), replace=True)
    sample_selection.to_excel('output_file.xlsx')