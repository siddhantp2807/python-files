from openpyxl import load_workbook
import pandas as pd
import json

class loadxl() :
    def __init__(self, filename):
        self.filename = filename
        pass
    def loadxlCols(self) :
        workbook = load_workbook(filename=self.filename)
        sheet = workbook.active

        list_with_values = []
        for cell in sheet[1] :
            list_with_values.append(cell.value)

        return list_with_values

    def dataframe(self) :
        cols = self.loadxlCols()
        data = pd.read_excel(self.filename)
        df = pd.DataFrame(data, columns = cols)

        return df

    def getXLData(self) :
        cols = self.loadxlCols()
        df = self.loadDataframe()
        finalArr = []
        for i in range(len(df)) :
            finalArr.append({cols[j] : df.iloc[i, j] for j in range(len(cols))} )
        return finalArr

    def writeToCsv(self) :
        fileName = self.filename.split('\\')[-1].split('.')[0] + '.csv'
        self.loadDataframe().to_csv(fileName, index=False)

    def writeToJSON(self) :
        fileName = self.filename.split('\\')[-1].split('.')[0] + '.json'
        data = self.loadValues()
        with open(fileName, 'w') as fileOpen :
            json.dump(data, fileOpen)
        


class loadJson() :
    def __init__(self, filename) :
        self.filename = filename

    def getJSONData(self) :
        with open(self.filename, 'r') as fileOpen :
            data = json.load(fileOpen)
        return data

    def dataframe(self) :
        return pd.DataFrame(self.getJSONData())

    def writeToxl(self) :

        fileName = self.filename.split(r'\\')[-1].split('.')[0] + '.xlsx'
        self.dataframe().to_excel(fileName, index=False)

        

    def writetoCSV(self) :
        fileName = self.filename.split(r'\\')[-1].split('.')[0] + '.csv'
        self.dataframe().to_csv(fileName, index=False)



class loadCsv() :
    def __init__(self, filename) :
        self.filename = filename

    def getCSVData(self) :
        df = pd.read_csv(self.filename)
        cols = df.columns
        finalArr = []
        for i in range(len(df)) :
            finalArr.append({cols[j] : df.iloc[i, j] for j in range(len(cols))} )
        return finalArr

    def dataframe(self) :
        return pd.DataFrame(self.getCSVData())

    def addToxl(self) :
        fileName = self.filename.split(r'\\')[-1].split('.')[0] + '.xlsx'

        df = pd.read_csv(self.filename)
        df.to_excel(fileName, index = False)
        
    def addToJSON(self) :
        fileName = self.filename.split(r'\\')[-1].split('.')[0] + '.json'

        with open(fileName, 'w') as file :
            json.dump(self.getCSVData(), file)




  
