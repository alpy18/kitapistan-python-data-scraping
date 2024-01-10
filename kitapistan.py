from selenium import webdriver
import xlsxwriter

browser = webdriver.Firefox()


workbook = xlsxwriter.Workbook('books.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({"bold": True})
worksheet.write("A1", "book_name", bold)
worksheet.write("B1", "image", bold)
worksheet.write("C1", "author", bold)
worksheet.write("D1", "publisher", bold)
worksheet.write("E1", "description", bold)
worksheet.write("F1", "pages", bold)
worksheet.write("G1", "category", bold)
worksheet.write("H1", "price", bold)
worksheet.write("I1", "best_seller", bold)

row = 1

book_urls = []

for i in range(1, 4):
    link = f"https://www.kitapyurdu.com/index.php?route=product/category&page={i}&filter_category_all=true&path=1&filter_in_stock=1&sort=purchased_365&order=DESC"
    browser.get(link)

    books = browser.find_elements_by_class_name("pr-img-link")
    
    for book in books:
        book_urls.append(book.get_attribute("href"))
        


for i in book_urls:
    browser.get(i)
    book_name = browser.find_element_by_css_selector(".pr_header__heading").text
    try:
        image = browser.find_element_by_id("js-book-cover").get_attribute("src")
    except:
        image = browser.find_element_by_class_name("js-jbox-book-cover").get_attribute("src")
    try:
        author = browser.find_element_by_css_selector(".pr_producers__manufacturer > div:nth-child(1) > a:nth-child(1)").text
    except:
        author = "Bilinmiyor"
    publisher = browser.find_element_by_css_selector(".pr_producers__publisher > div:nth-child(1) > a:nth-child(1)").text
    description = browser.find_element_by_class_name("info__text").text
    pages = browser.find_element_by_class_name("pr_attributes").text
    pages = pages.split("\n")
    pages = [i for i in pages if "Sayfa Sayısı" in i]
    try:
        pages = pages[0][14:]
    except:
        pages = "Bilinmiyor"
    cok_satan = "1"
    try:
        category = browser.find_element_by_xpath("/html/body/div[5]/div/div/div[8]/div/div[2]/div[2]/div[1]/div[2]/div[2]/div/ul/li[1]/a/span[2]").text
    except:
        category = "Diğer"
    price = browser.find_element_by_class_name("price__item").text

    worksheet.write(row, 0, row-1)
    worksheet.write(row, 0, book_name)
    worksheet.write(row, 1, image)
    worksheet.write(row, 2, author)
    worksheet.write(row, 3, publisher)
    worksheet.write(row, 4, description)
    worksheet.write(row, 5, pages)
    worksheet.write(row, 6, category)
    worksheet.write(row, 7, price)
    worksheet.write(row, 8, cok_satan)
    row += 1
    



cat_list = ["https://www.kitapyurdu.com/index.php?route=product/category&page=4&filter_category_all=true&path=1_236&filter_in_stock=1&sort=purchased_365&order=DESC",
            "https://www.kitapyurdu.com/index.php?route=product/category&page=4&filter_category_all=true&path=1_2&filter_in_stock=1&sort=purchased_365&order=DESC",
            "https://www.kitapyurdu.com/index.php?route=product/category&page=4&filter_category_all=true&path=1_128&filter_in_stock=1&sort=purchased_365&order=DESC",
            "https://www.kitapyurdu.com/index.php?route=product/category&page=4&filter_category_all=true&path=1_359&filter_in_stock=1&sort=purchased_365&order=DESC",
            "https://www.kitapyurdu.com/index.php?route=product/category&page=4&filter_category_all=true&path=1_424&filter_in_stock=1&sort=purchased_365&order=DESC",
            "https://www.kitapyurdu.com/index.php?route=product/category&page=4&path=341&filter_in_stock=1",
            "https://www.kitapyurdu.com/index.php?route=product/category&page=4&filter_category_all=true&path=1_161&filter_in_stock=1&sort=purchased_365&order=DESC",
            "https://www.kitapyurdu.com/index.php?route=product/category&page=4&filter_category_all=true&path=1_87&filter_in_stock=1&sort=purchased_365&order=DESC",
            "https://www.kitapyurdu.com/index.php?route=product/category&page=4&filter_category_all=true&path=1_41&filter_in_stock=1&sort=purchased_365&order=DESC"]



book_urls2 = []

categories = [
  "Bilim ve Mühendislik",
  "Çocuk Kitapları",
  "Edebiyat",
  "Eğitim",
  "Felsefe",
  "Kişisel Gelişim",
  "Müzik",
  "Psikoloji",
  "Tarih"
]

for i in cat_list:
    browser.get(i)

    books = browser.find_elements_by_class_name("pr-img-link")
    
    for book in books:
        book_urls2.append(book.get_attribute("href"))
        

county = 0
indexy = 0

for i in range(len(book_urls2)):
    if (i % 20 != 18) and (i % 20 != 19):
        browser.get(book_urls2[i])
        book_name = browser.find_element_by_css_selector(".pr_header__heading").text
        try:
            image = browser.find_element_by_id("js-book-cover").get_attribute("src")
        except:
            image = browser.find_element_by_class_name("js-jbox-book-cover").get_attribute("src")
        try:
            author = browser.find_element_by_css_selector(".pr_producers__manufacturer > div:nth-child(1) > a:nth-child(1)").text
        except:
            author = "Bilinmiyor"
        publisher = browser.find_element_by_css_selector(".pr_producers__publisher > div:nth-child(1) > a:nth-child(1)").text
        description = browser.find_element_by_class_name("info__text").text
        pages = browser.find_element_by_class_name("pr_attributes").text
        pages = pages.split("\n")
        pages = [i for i in pages if "Sayfa Sayısı" in i]
        try:
            pages = pages[0][14:]
        except:
            pages = "Bilinmiyor"
        cok_satan = "0"
        category = categories[indexy]
        price = browser.find_element_by_class_name("price__item").text

        worksheet.write(row, 0, row-1)
        worksheet.write(row, 0, book_name)
        worksheet.write(row, 1, image)
        worksheet.write(row, 2, author)
        worksheet.write(row, 3, publisher)
        worksheet.write(row, 4, description)
        worksheet.write(row, 5, pages)
        worksheet.write(row, 6, category)
        worksheet.write(row, 7, price)
        worksheet.write(row, 8, cok_satan)
        row += 1

        county += 1
        
        if county == 18:
            indexy += 1
            county = 0





workbook.close()
browser.close()


