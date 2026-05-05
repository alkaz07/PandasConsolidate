import pandas as pd
import matplotlib.pyplot as plt

cities = ["Гадюкино", "Нью-Васюки", "Мирный"]   #исходный список названий городов
dataframes = []
for city in cities:
    filename = city + '.xlsx'
    try:
        df = pd.read_excel(filename)
        df['Город'] = city
        dataframes.append(df)
    except FileNotFoundError:
        print(f'файл с таким именем {filename} не найден в указанном месте')

df_all = pd.concat(dataframes, ignore_index=True)

print(df_all)
df_sums_by_position = df_all.groupby('должность')['выплата'].sum()
print(df_sums_by_position)
df_sums_by_position.plot(kind = 'pie')

plt.show()  #обязательно надо, чтобы увидеть картинку в пайчарме