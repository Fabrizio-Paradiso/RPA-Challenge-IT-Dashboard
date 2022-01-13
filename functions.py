from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem

class Functions:
    list_amounts = {}
    list_individual_investments = {}

    def init_browser(self,url):
        self.driver = Selenium()
        self.files = FileSystem()
        self.excel = Files()

        self.driver.set_download_directory(self.files.absolute_path("output/"))
        self.driver.open_available_browser(url)
        self.driver.maximize_browser_window()

    def close_browsers(self):
        self.driver.close_all_browsers()

    def click_dive_in_button(self):
        self.driver.click_element_when_visible("class:btn-lg-2x")

    def get_agencies_amounts(self):
        AMOUNTS = []
        AGENCIES = []
        try:
            self.driver.wait_until_page_contains_element('//*[@id="agency-tiles-widget"]//span[@class=" h1 w900"]')
        except:
            print("Amounts agencies were not found")

        for element in self.driver.find_elements('//*[@id="agency-tiles-widget"]//span[@class="h4 w200"]'):
            AGENCIES.append(element.text)
        
        for element in self.driver.find_elements('//*[@id="agency-tiles-widget"]//span[@class=" h1 w900"]'):
            AMOUNTS.append(element.text)

        self.list_amounts = {'Agency': AGENCIES, 'Amount': AMOUNTS}
        self.write_agencies_amounts_to_excel_sheet("output/file.xlsx",self.list_amounts,"Agencies")

    def write_agencies_amounts_to_excel_sheet(self,excel_path,content,sheet_name):
        file = self.excel.create_workbook(excel_path)
        file.append_worksheet("Sheet",content)
        file.rename_worksheet(sheet_name, "Sheet")
        file.save()
        
    def get_agency_individual_investments(self,keyword_index):
        INVESTMENT_TITLE = []
        INDIVIDUAL_INVESTMENTS = []

        try:
            for index,element in enumerate(self.driver.find_elements('//*[@id="agency-tiles-widget"]//a[@class="btn btn-default btn-sm"]'),start=1):
                if(index==keyword_index):
                    self.driver.go_to(self.driver.get_element_attribute(element,"href"))
        except:
            print("An error ocurred in one of the elements")
        
        try: 
            self.driver.wait_until_page_contains_element('//*[@id="investments-table-object_length"]/label/select',timeout=15)
            self.driver.select_from_list_by_label('//*[@id="investments-table-object_length"]/label/select',"All")
        except:
            print("Selector was not found")

        self.driver.wait_until_element_is_not_visible('//*[@id="investments-table-object_paginate"]/span/a[2]',timeout=15)

        try:
            for element in self.driver.find_elements('//*[@id="investments-table-object"]//td[@class=" left"]'):
                INVESTMENT_TITLE.append(element.text)
        except:
            print("An error ocurred in one of the elements")

        try:
            for element in self.driver.find_elements('css:td.right'):
                INDIVIDUAL_INVESTMENTS.append(element.text)
        except:
            print("An error ocurred in one of the elements")

        self.list_individual_investments = {'Investment Title': INVESTMENT_TITLE, 'Total': INDIVIDUAL_INVESTMENTS}
        self.write_individual_investments_to_excel_sheet("output/file.xlsx",self.list_individual_investments,"Individual Investments")

    def write_individual_investments_to_excel_sheet(self,excel_path,content,sheet_name):
        file = self.excel.open_workbook(excel_path)
        file.create_worksheet(sheet_name)
        file.append_worksheet(sheet_name,content)
        file.save()

    def download_business_case_pdf(self):
        links = []
        try:
            for element in self.driver.find_elements('css:td.sorting_2 a'):
                links.append(self.driver.get_element_attribute(element,"href"))
        except:
            print("There's no link existing for this agency")
            
        print(len(links))
        for link in links:

            self.driver.go_to(link)
            self.driver.wait_until_element_is_visible('//*[@id="business-case-pdf"]/a')
            self.driver.click_element('//*[@id="business-case-pdf"]/a')

            while True:
                try:
                    if self.driver.find_element('//*[@id="business-case-pdf"]/span'):
                        pass
                    else:
                        break
                except:
                    if self.driver.find_element('//*[@id="business-case-pdf"]//a[@aria-busy="false"]'):
                        break

        try:
            for file in self.files.list_files_in_directory(self.files.absolute_path("output/")):
                if(self.files.get_file_extension(file) == '.crdownload' or self.files.get_file_extension(file) == '.tmp'):
                    self.files.wait_until_removed(self.files.join_path(self.files.absolute_path("output/"),file),timeout=25)
        except:
            print("Download coudn't be finished")
    

        