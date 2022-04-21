import pandas as pd
import iteround


class DataParser:
    def __init__(self):
        self.data = None

    def load_transform_data(self, data_filename: str, exchange_filename: str):
        """
        Method that takes two Excel file names, read them into pandas DataFrames, then transform the data into desired
        output, and save it into class variable 'data'.
        """
        # Load Excel files
        df = pd.read_excel(io=data_filename)
        exchange = pd.read_excel(io=exchange_filename)

        # Get annual revenue of each company in USD, rename the column and round it to two decimal places
        df = df.groupby('Company').mean().reset_index()
        df = df.rename(columns={'Revenue': 'Revenue (USD)'})
        df['Revenue (USD)'] = round(df['Revenue (USD)'], 2)

        # Get annual revenue in CZK and share of the total revenue in %, both rounded to two decimals places
        df['Revenue (Local)'] = round(df['Revenue (USD)'] * exchange.loc[exchange['Country'] == 'Czech Republic']
                                      .iloc[0]['Annual Rate'], 2)
        # I decided to use iteround library for rounding share, so the sum stays the same (100)
        df['Share'] = iteround.saferound(((df['Revenue (USD)'] / df['Revenue (USD)'].sum()) * 100).tolist(), 2)

        # Add total statistics
        df.loc[len(df.index)] = ['Total', df['Revenue (USD)'].sum(), df['Revenue (Local)'].sum(), df['Share'].sum()]

        # Add corresponding prefixes and suffixes to column values
        df['Revenue (USD)'] = '$' + df['Revenue (USD)'].astype(str)
        df['Revenue (Local)'] = 'CZK ' + df['Revenue (Local)'].astype(str)
        df['Share'] = df['Share'].astype(str) + ' %'

        # Save transformed data to class variable
        self.data = df

    def get_company_revenue(self, company_name: str) -> str:
        # Get row with given company, then print its revenue in USD and its share
        company_row = self.data.loc[self.data['Company'] == company_name]
        return f"Company {company_name} has revenue {company_row.iloc[0]['Revenue (USD)']} and share " \
               f"{company_row.iloc[0]['Share']}"

    def get_company_row_number(self, company_name: str) -> str:
        # Get row index of given company and print it (also add 1 to it because indices start on 0) and print the result
        row_index = self.data.index[self.data['Company'] == company_name]
        return f'Company {company_name} is located on a row number {row_index[0] + 1}'

    def sort_by_company(self, ascending: bool) -> pd.DataFrame:
        # Sort data by company name and print the result
        self.data = self.data.sort_values(by='Company', ascending=ascending)
        return self.data.head()

    def sort_by_revenue(self, ascending: bool) -> pd.DataFrame:
        # Sort data by revenue (dollar sign needs to be removed before sorting and then added back) and print the result
        self.data['Revenue (USD)'] = self.data['Revenue (USD)'].str[1:].astype(float)
        self.data = self.data.sort_values(by='Revenue (USD)', ascending=ascending)
        self.data['Revenue (USD)'] = '$' + self.data['Revenue (USD)'].astype(str)
        return self.data.head()

    def export_html(self, filename: str):
        html_file = open(filename, 'w')
        html_file.write(self.data.to_html(index=False))

    def export_excel(self, filename: str):
        self.data.to_excel(filename, index=False)

    def export_csv(self, filename: str):
        self.data.to_csv(filename, index=False)


if __name__ == "__main__":
    dp = DataParser()
    dp.load_transform_data(data_filename="data.xlsx", exchange_filename="exchange.xlsx")

    print(dp.get_company_revenue(company_name='Apple'))
    print(dp.get_company_row_number(company_name='Apple'))

    print(dp.sort_by_company(ascending=True))
    print(dp.sort_by_revenue(ascending=False))

    dp.export_html(filename='output.html')
    dp.export_excel(filename='output.xlsx')
    dp.export_csv(filename='output.csv')
