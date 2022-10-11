import pandas as pd
import matplotlib.pyplot as plt
# 
# Requirments
# 
# [x] Als Danny kan ik een excel inladen en die weergeven in het dashboard
# [x] Als Danny wil ik weten hoeveel sales ieder team draait
# [x] Als Danny wil ik weten welk sales team onder gemiddeld scoort
# [ ] Als Danny wil ik inzichtelijk hebben welke sales team regio het beste draait
# 
# https://datatofish.com/read_excel/


class NuNu():


    def __init__(self):
        self.sheetLoader()
        self.salesTeamPluiser()

    def loadSheet(self, sheet):        
        # Als Danny kan ik een excel inladen en die weergeven in het dashboard
        return pd.read_excel (r'sales_data.xlsx', sheet_name=sheet)
        
    def sheetLoader(self):
        self.sheets = {
            'orders' : self.loadSheet('Sales Orders Sheet'),
            'teams' : self.loadSheet('Sales Team Sheet'),
            # 'regions' : self.loadSheet('Regions Sheet'),
            # 'products' : self.loadSheet('Products Sheet'),
            # 'store_locations' : self.loadSheet('Store Locations Sheet'),
            # 'customers' : self.loadSheet('Customers Sheet'),
        }


    def fetchTeam(self, id):
                
        dataset = pd.DataFrame(self.sheets['teams'])                
        return dataset.loc[dataset['_SalesTeamID'] == id]['Sales Team'].values[0]


    def fetchRegion(self, id):
                
        dataset = pd.DataFrame(self.sheets['teams'])                
        dataset.loc[dataset['_SalesTeamID'] == id]['Region'].values[0]


    def salesTeamPluiser(self):
        
        # Als Danny wil ik weten hoeveel sales ieder team draait
        dataset = pd.DataFrame(self.sheets['orders'], columns= ['_SalesTeamID','Order Quantity', 'Unit Price', 'Unit Cost', 'Discount Applied'])
        self.salesTeam = {}
        for index, row in dataset.iterrows():
            some_math_result = row['Unit Cost'] - ((row['Order Quantity'] * row['Unit Price']) * row['Discount Applied'])
            if row['_SalesTeamID'] in self.salesTeam:
                self.salesTeam[row['_SalesTeamID']] += int(some_math_result)
            else:
                self.salesTeam[row['_SalesTeamID']] = int(some_math_result)


        # Als Danny wil ik weten welk sales team onder gemiddeld scoort
        self.salesTeamAverage = sum(self.salesTeam.values()) / len(self.salesTeam)
        self.salesTeam = dict(sorted(self.salesTeam.items(), key=lambda item: item[1], reverse=True))


        names = []
        data = []
        for row in self.salesTeam:
            data.append(self.salesTeam[row])
            names.append(self.fetchTeam(row)) 

        plt.bar(names, data)
        plt.ylabel('OLA')
        plt.show()



        # for row in self.salesTeam:

            
        #     if (self.salesTeam[row] < self.salesTeamAverage):
        #         print(f"\033[91m {self.salesTeam[row]}  -   {self.fetchTeam(row)}")
        #     else:
        #         print(f"\033[92m {self.salesTeam[row]}  -   {self.fetchTeam(row)}")
                
                        
        # #  Working for order quantity
        # dataset = dataset.groupby('_SalesTeamID').sum().sort_values('Order Quantity', ascending=False)
        # for index, row in dataset.iterrows():
        #     print(index, self.fetchTeam(index), row['Order Quantity'])

        
mainframe = NuNu()
