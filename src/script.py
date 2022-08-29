from .xlanalysis import xl2pddf



df = xl2pddf('data/bp-stats-review-2019-all-data.xlsx', 1)

print(df)