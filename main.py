import os
import re
import time
import logging
from datetime import timedelta

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.PDF import PDF


browser_lib = Selenium()
excel_file = Files()
pdf = PDF()


def open_the_website(url):
    download_path = os.path.abspath("output")

    browser_lib.set_download_directory(download_path)
    browser_lib.open_available_browser(url=url)


def click_button():
    browser_lib.click_element("link:DIVE IN")
    browser_lib.wait_until_element_is_visible('id:agency-tiles-container')


def get_agency_list():
    container_element = browser_lib.find_element('id:agency-tiles-container')
    container_text = container_element.text.split('\n')

    result_list = list(zip(container_text[::4], container_text[2::4]))
    result_list.insert(0, ("Agency name", "Spending"))

    return result_list


def write_to_excel(sheetname, bookname, content, method):
    if not os.path.exists("output"):
        os.mkdir("output")
    book_path = os.path.join("output", bookname)

    if method == "write":
        excel_file.create_workbook(book_path)
    elif method == "append":
        excel_file.open_workbook(book_path)

    if not excel_file.worksheet_exists(sheetname):
        excel_file.create_worksheet(name=sheetname)

    excel_file.set_active_worksheet(sheetname)
    excel_file.append_rows_to_worksheet(content)
    excel_file.save_workbook()


def select_department(name):
    selector = "id:agency-tiles-widget >> class:tuck-5"
    department_element = [i for i in browser_lib.find_elements(selector) if name in i.text][0]

    browser_lib.click_element(department_element)
    browser_lib.wait_until_element_is_visible("id:investments-table-widget", timeout=timedelta(seconds=20))


def scrape_table(headers):
    selection = browser_lib.find_element("name:investments-table-object_length")
    selection.click()
    browser_lib.click_element(selection.find_element_by_xpath("//option[@value=-1]"))

    condition = "return document.getElementById('investments-table-object_last').classList.contains('disabled')"
    browser_lib.wait_for_condition(condition=condition, timeout=timedelta(seconds=20))

    table = browser_lib.find_element('id:investments-table-object >> tag:tbody')
    table_data = [headers]
    links = []
    equalities = {}  # dict with "UII": "Name of investment"

    for row in table.find_elements_by_tag_name("tr"):
        cols = row.find_elements_by_tag_name("td")
        data_set = []

        for data in cols:
            data_set.append(data.text)
            a_element = data.find_elements_by_tag_name("a")
            if a_element:
                links.append(a_element[0].get_attribute("href"))

        table_data.append(data_set)
        equalities[data_set[0]] = data_set[2]

    write_to_excel(sheetname="Individual Investments", bookname="spending.xlsx", content=table_data, method="append")

    for link in links:
        file = download_file(link)
        compare_values(equalities, file)


def download_file(link):
    browser_lib.go_to(link)
    element = 'id:business-case-pdf >> tag:a'

    browser_lib.wait_until_element_is_visible(element, timeout=timedelta(seconds=20))
    browser_lib.click_element(element)
    browser_lib.assign_id_to_element(locator=element, id="USERFLAG")

    condition = "return document.getElementById('USERFLAG').getAttribute('aria-busy') == \"false\""
    browser_lib.wait_for_condition(condition=condition, timeout=timedelta(seconds=20))

    time.sleep(3)  # wait till file is downloaded

    return get_recent_file("output")


def get_recent_file(path):
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files]
    return max(paths, key=os.path.getctime)


def compare_values(equalities, file):
    text = pdf.get_text_from_pdf(file)[1]
    investment = re.search(r"1\. Name of this Investment: (.*)2\.", text).group(1)
    uii = re.search(r"2\. Unique Investment Identifier \(UII\): (.*)Section B", text).group(1)

    if equalities[uii] == investment:
        logging.info(uii, investment, "EQUAL")
    else:
        logging.info(uii, investment, "NOT EQUAL")


def clean_dir(folder):
    path = os.path.abspath(folder)
    for item in os.listdir(path):
        if (not item.endswith(".xlsx")) and (not item.endswith(".pdf")):
            os.remove(os.path.join(path, item))


def main():
    try:
        open_the_website("https://itdashboard.gov/drupal/")
        click_button()
        write_to_excel("Agencies", "spending.xlsx", content=get_agency_list(), method="write")
        select_department("Department of the Interior")
        scrape_table(["UII", "Bureau", "Investment Title", "Total FY2021 Spending ($M)", "Type", "CIO Rating", "# of Projects"])

    finally:
        browser_lib.close_all_browsers()
        clean_dir("output")


if __name__ == "__main__":
    main()

