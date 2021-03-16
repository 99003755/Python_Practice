import pandas as pd

raw_data = {'first_name': ['Sheldon', 'Raj', 'Leonard', 'Howard', 'Amy'],
            'last_name': ['Copper', 'Koothrappali', 'Hofstadter', 'Wolowitz', 'Fowler'],
            'age': [42, 38, 36, 41, 35],
            'Comedy_Score': [9, 7, 8, 8, 5],
            'Rating_Score': [25, 25, 49, 62, 70]}

df = pd.DataFrame(raw_data, columns=['first_name', 'last_name', 'age',
                                     'Comedy_Score', 'Rating_Score'])
print(df)

