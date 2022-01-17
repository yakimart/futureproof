from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
import os

browser_lib = Selenium()
excel_file = Files()


def open_the_website(url):
    download_path = os.path.abspath("output")
    pref = {"download.default_directory": download_path}
    browser_lib.open_available_browser(url=url, preferences=pref)


def click_button(button_text, container):
    a_element = f'//*[contains(text(), "{button_text}")]'
    browser_lib.click_element(a_element)
    browser_lib.wait_until_element_is_visible(f'//*[@id="{container}"]')


def get_agency_list():
    container_element = browser_lib.find_element('//*[@id="agency-tiles-container"]')
    container_text = container_element.text.split('\n')
    result_list = list(zip(container_text[::4], container_text[2::4]))
    result_list.insert(0, ("Agency name", "Spending"))
    return result_list


def write_to_excel(sheetname, output):
    if os.path.exists(output):
        excel_file.open_workbook(output)
    else:
        excel_file.create_workbook(output)

    if not excel_file.worksheet_exists(sheetname): excel_file.create_worksheet(name=sheetname)
    excel_file.set_active_worksheet(sheetname)
    excel_file.append_rows_to_worksheet(content=get_agency_list())
    excel_file.save_workbook()


def select_department(name):
    department_element = [i for i in browser_lib.find_elements('css:#agency-tiles-widget >> css:.tuck-5') if name in i.text][0]
    browser_lib.click_element(department_element)
    browser_lib.wait_until_element_is_visible('//*[@id="investments-table-widget"]/div', timeout=15)


def scrape_table():
    selection = browser_lib.find_element('//*[@id="investments-table-object_length"]/label/select')
    selection.click()
    browser_lib.click_element(selection.find_element_by_xpath("//option[@value=-1]"))

    condition = "return document.getElementById('investments-table-object_last').classList.contains('disabled')"
    browser_lib.wait_for_condition(condition=condition, timeout=20)

    table = (browser_lib.find_element('//*[@id="investments-table-object"]/tbody'))
    links = [link.get_attribute("href") for link in table.find_elements_by_tag_name("a")]
    table_data = [["UII", "Bureau", "Investment Title", "Total FY2021 Spending ($M)", "Type", "CIO Rating", "# of Projects"]]

    for row in table.find_elements_by_tag_name("tr"):
        cols = row.find_elements_by_tag_name("td")
        data_set = []
        for data in cols:
            data_set.append(data.text)
        table_data.append(data_set)

    excel_file.open_workbook("output/spending.xlsx")
    if not excel_file.worksheet_exists("Individual Investments"): excel_file.create_worksheet(name="Individual Investments")
    excel_file.set_active_worksheet("Individual Investments")
    excel_file.append_rows_to_worksheet(content=table_data)
    excel_file.save_workbook()

    for link in links:
        browser_lib.go_to(link)
        element = '//*[@id="business-case-pdf"]/a'
        browser_lib.wait_until_element_is_visible(element, timeout=10)
        browser_lib.click_element(element)
        browser_lib.assign_id_to_element(locator=element, id="USERFLAG")
        condition = "return document.getElementById('USERFLAG').getAttribute('aria-busy') == \"false\""
        browser_lib.wait_for_condition(condition=condition, timeout=10)


def main():
    try:
        open_the_website("https://itdashboard.gov/drupal/")
        click_button("DIVE IN", "agency-tiles-container")
        write_to_excel("Agencies", "output/spending.xlsx")
        select_department("Department of the Interior")
        scrape_table()

    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()

