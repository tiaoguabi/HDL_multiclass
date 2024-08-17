# %%
import os
import pandas as pd
import warnings
from tableone import TableOne
from scipy import stats

warnings.filterwarnings("ignore")

# %%
work_dir = 'E:/Disk E/Grand Blue/Research and study/HDL_multiclass'
os.chdir(work_dir)

os.makedirs('./tables', exist_ok=True)
# %%
df = pd.read_excel('data/TertileClass_Gensini_data.xlsx')

# %%
list1 = df.columns.to_list()
# %%
# describes the distribution of variables
columns_to_encode = ['hypertension', 'diabetes', 'stroke', 'kidney disease', 'Thyroid Dysfunction', 'COPD']

df_encoded = df.copy()

for column in columns_to_encode:
    df_encoded[column] = df_encoded[column].replace({1: 'Yes', 0: 'No'})

df_encoded['sex'] = df_encoded['sex'].replace({1: 'Male', 0: 'Female'})
df_encoded['Gensini_tertile'] = df_encoded['Gensini_tertile'].replace({0: 'mild', 1: 'moderate', 2: 'severe'})

columns_describe = ['age', 'sex', 'hypertension', 'diabetes', 'stroke', 'kidney disease', 'Thyroid Dysfunction', 'COPD',
                    'TC', 'TG', 'LDL-C', 'HDL-C',
                    'HDL-2b', 'HDL-3', 'Gensini_total_Score', 'Gensini_tertile', 'Gensini_tertile_label']

description = df_encoded[columns_describe].describe(include='all').T

description = description.applymap(lambda x: round(x, 1) if isinstance(x, (int, float)) else x)

output_file = 'tables/columns_description.xlsx'
description.to_excel(output_file)

# %%
categorical = ['sex', 'hypertension', 'diabetes', 'stroke', 'kidney disease', 'Thyroid Dysfunction', 'COPD',
               'Gensini_tertile', 'Gensini_tertile_label']

list_var = [k for k in columns_describe if k not in categorical]
nonnormal = []

for n in list_var:
    N, p = stats.normaltest(df_encoded[n])
    print(N, p)
    if p < 0.05:
        print(n, 'Non-normal Distribution')
        nonnormal.append(n)
    else:
        print(n, 'Normal Distribution')
    print('\n')

mytable = TableOne(df_encoded, columns=columns_describe, nonnormal=nonnormal, categorical=categorical,
                   groupby='Gensini_tertile',
                   normal_test=True)
print(mytable.tabulate(tablefmt="github"))

mytable.to_excel('./tables/table1.xlsx')
