import pandas as pd

df = pd.DataFrame({'a': [1, 2, 3], 'b': [4, 5, 6]})
df_1 = pd.DataFrame({'b': [4]})

df_concat = pd.concat([df, df_1], axis=1)
print(df_concat)