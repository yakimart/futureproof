import os
from datetime import timedelta
from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium


browser_lib = Selenium()
excel_file = Files()


def open_the_website(url):
    download_path = os.path.abspath("output")
    pref = {"download.default_directory": download_path}
    browser_lib.open_available_browser(url=url, preferences=pref)
    browser_lib.set_download_directory(download_path)
    # special method <set_download_directory> has no effect


def click_button():
    browser_lib.click_element("link:DIVE IN")
    browser_lib.wait_until_element_is_visible('id:agency-tiles-container')


def get_agency_list():
    container_element = browser_lib.find_element('id:agency-tiles-container')
    container_text = container_element.text.split('\n')
    result_list = list(zip(container_text[::4], container_text[2::4]))
    result_list.insert(0, ("Agency name", "Spending"))
    return result_list


def write_to_excel(sheetname, bookname, content):
    if not os.path.exists("output"):
        os.mkdir("output")
    book_path = os.path.join("output", bookname)
    if os.path.exists(book_path):
        excel_file.open_workbook(book_path)
    else:
        excel_file.create_workbook(book_path)

    if not excel_file.worksheet_exists(sheetname):
        excel_file.create_worksheet(name=sheetname)
    excel_file.set_active_worksheet(sheetname)
    excel_file.append_rows_to_worksheet(content)
    excel_file.save_workbook()


def select_department(name):
    selector = "id:agency-tiles-widget >> class:tuck-5"
    department_element = [i for i in browser_lib.find_elements(selector) if name in i.text][0]
    browser_lib.click_element(department_element)
    browser_lib.wait_until_element_is_visible("id:investments-table-widget", timeout=timedelta(seconds=15))


def scrape_table(headers):
    selection = browser_lib.find_element("name:investments-table-object_length")
    selection.click()
    browser_lib.click_element(selection.find_element_by_xpath("//option[@value=-1]"))

    condition = "return document.getElementById('investments-table-object_last').classList.contains('disabled')"
    browser_lib.wait_for_condition(condition=condition, timeout=timedelta(seconds=20))

    table = browser_lib.find_element('//*[@id="investments-table-object"]/tbody')
    links = [link.get_attribute("href") for link in table.find_elements_by_tag_name("a")]
    table_data = [headers]

    for row in table.find_elements_by_tag_name("tr"):
        cols = row.find_elements_by_tag_name("td")
        data_set = []
        for data in cols:
            data_set.append(data.text)
        table_data.append(data_set)

    write_to_excel(sheetname="Individual Investments", bookname="spending.xlsx", content=table_data)

    for link in links:
        browser_lib.go_to(link)
        element = '//*[@id="business-case-pdf"]/a'
        browser_lib.wait_until_element_is_visible(element, timeout=timedelta(seconds=10))
        browser_lib.click_element(element)
        browser_lib.assign_id_to_element(locator=element, id="USERFLAG")
        condition = "return document.getElementById('USERFLAG').getAttribute('aria-busy') == \"false\""
        browser_lib.wait_for_condition(condition=condition, timeout=timedelta(seconds=10))


def main():
    try:
        open_the_website("https://itdashboard.gov/drupal/")
        click_button()
        write_to_excel("Agencies", "spending.xlsx", content=get_agency_list())
        select_department("Department of the Interior")
        scrape_table(["UII", "Bureau", "Investment Title", "Total FY2021 Spending ($M)", "Type", "CIO Rating", "# of Projects"])

    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()

# The task below positioned as bonus (optional).


# 'Extract data from PDF. You need to get the data from Section A in each PDF.
# Then compare the value "Name of this Investment" with the column "Investment Title",
# and the value "Unique Investment Identifier (UII)" with the column "UII"'


# It is not mentioned what exactly the program should do after comparing. What should happen if, for example,
# the values in pdf and table are different? 
