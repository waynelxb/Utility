import unittest
from selenium import webdriver
class WebTable:
    def __init__(self, webtable):
       self.table = webtable

    def get_row_count(self):
      return len(self.table.find_elements_by_tag_name("tr")) - 1

    def get_column_count(self):
        return len(self.table.find_elements_by_xpath("//tr[2]/td"))

    def get_table_size(self):
        return {"rows": self.get_row_count(),
                "columns": self.get_column_count()}

    def row_data(self, row_number):
        if(row_number == 0):
            raise Exception("Row number starts from 1")

        row_number = row_number + 1
        row = self.table.find_elements_by_xpath("//tr["+str(row_number)+"]/td")
        rData = []
        for webElement in row :
            rData.append(webElement.text)

        return rData

    def column_data(self, column_number):
        col = self.table.find_elements_by_xpath("//tr/td["+str(column_number)+"]")
        rData = []
        for webElement in col :
            rData.append(webElement.text)
        return rData

    def get_all_data(self):
        # get number of rows
        noOfRows = len(self.table.find_elements_by_xpath("//tr")) -1
        # get number of columns
        noOfColumns = len(self.table.find_elements_by_xpath("//tr[2]/td"))
        allData = []
        # iterate over the rows, to ignore the headers we have started the i with '1'
        for i in range(2, noOfRows):
            # reset the row data every time
            ro = []
            # iterate over columns
            for j in range(1, noOfColumns) :
                # get text from the i th row and j th column
                ro.append(self.table.find_element_by_xpath("//tr["+str(i)+"]/td["+str(j)+"]").text)

            # add the row data to allData of the self.table
            allData.append(ro)

        return allData

    def presence_of_data(self, data):

        # verify the data by getting the size of the element matches based on the text/data passed
        dataSize = len(self.table.find_elements_by_xpath("//td[normalize-space(text())='"+data+"']"))
        presence = False
        if(dataSize > 0):
            presence = True
        return presence

    def get_cell_data(self, row_number, column_number):
        if(row_number == 0):
            raise Exception("Row number starts from 1")

        row_number = row_number+1
        cellData = self.table.find_element_by_xpath("//tr["+str(row_number)+"]/td["+str(column_number)+"]").text
        return cellData



driver = webdriver.Chrome(executable_path="D:/Projects/GitHub/UtilityRepository/BookTennisCourt/chromedriver.exe"   )
driver.implicitly_wait(30)

driver.get("https://chercher.tech/practice/table")
w = WebTable(driver.find_element_by_xpath("//table[@id='webtable']"))

print("No of rows : ", w.get_row_count())
print("------------------------------------")
print("No of cols : ", w.get_column_count())
print("------------------------------------")
print("Table size : ", w.get_table_size())
print("------------------------------------")
print("First row data : ", w.row_data(1))
print("------------------------------------")
print("First column data : ", w.column_data(1))
print("------------------------------------")
print("All table data : ", w.get_all_data())
print("------------------------------------")
print("presence of data : ", w.presence_of_data("Chercher.tech"))
print("------------------------------------")
print("Get data from Cell : ", w.get_cell_data(2, 2))

