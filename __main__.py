from functions import Functions
from keywords import*

def main():
    function = Functions ()
    try:
        function.init_browser(url="https://itdashboard.gov/")
        function.click_dive_in_button()
        function.get_agencies_amounts()
        function.get_agency_individual_investments(keyword_index)
        function.download_business_case_pdf()
    finally:
        function.close_browsers()

if __name__ == "__main__":
    main()