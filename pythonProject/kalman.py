import pandas as pd
import seaborn as sns
from statsmodels.tsa.seasonal import MSTL
from statsmodels.tsa.seasonal import STL
from matplotlib import pyplot as plt

# sns.set_context("talk")


# res = STL(df_total_long.loc[cond, 'Количество'].interpolate(method="linear"), seasonal=25).fit()
res = MSTL(cheburek.loc[cond, 'Количество'].interpolate(method="linear"), periods=(24, 24*7)).fit()


plt.rc("figure", figsize=(10, 10))
plt.rc("font", size=5)
res.plot()

seasonal_component = res.seasonal
seasonal_component.head()

df_deseasonalised = cheburek.loc[cond, 'Количество'] - seasonal_component['seasonal_24'] - seasonal_component['seasonal_168']
df_deseasonalised_imputed = df_deseasonalised.interpolate(method="linear")
df_imputed = df_deseasonalised_imputed + seasonal_component['seasonal_24'] + seasonal_component['seasonal_168']
df_imputed = df_imputed.to_frame().rename(columns={0: "Количество"})
ax = df_imputed.plot(linestyle="-", marker=".", figsize=[10, 5], legend=None)
ax = df_imputed[cheburek.loc[cond, 'Количество'].isnull()].plot(ax=ax, legend=None, marker=".", color="r")

ax.set_title("Retail Sales with imputed data")
ax.set_ylabel("Retail Sales")
ax.set_xlabel("Time")
# Apply the linear interpolation method
# Note: If the time intervals between rows are not uniform then
# the method should be set as 'time' to achieve a linear fit.
df_imputed = cheburek.loc[cond, 'Количество'].interpolate(method="spline", order=2)

# Plot the imputed time series
ax = df_imputed.plot(linestyle="-", marker=".", figsize=[10, 5], legend=None)
df_imputed[cheburek.loc[cond, 'Количество'].isnull()].plot(ax=ax, legend=None, marker=".", color="r")

ax.set_title("Retail Sales with imputed data")
ax.set_ylabel("Retail Sales")
ax.set_xlabel("Time")