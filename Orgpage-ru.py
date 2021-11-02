import pymysql.cursors
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

num_=0
_host = 'localhost'
_user = 'root'
_password = 'cfitymrf'
_db = 'scr2'
_charset = 'utf8'
debug = True
WebDriverWaitInSec = 30

def clean(txt):
    if txt == None:
        txt = ""
    txt = str.strip(txt)
    txt = str.replace(txt, "\"\"", "")
    txt = str.replace(txt, "\"", "")
    txt = str.replace(txt, "'", "")
    txt = txt.replace('\n', ' ')
    txt = txt.replace('\r', '')
    return txt


def getconnect():
    connect = pymysql.connect(host = _host,
                              user = _user,
                              password = _password,
                              db = _db,
                              charset = _charset,
                              cursorclass = pymysql.cursors.DictCursor)
    return connect


def disconnect(connection):
    connection.close()


def ins(connection, table, ischeckdupl, flddupl, valdupl, **kwargs):
    id = 0
    global debug
    try:
        fields = ""
        vals = ""
        if ischeckdupl == True:
            with connection.cursor() as cur:
                sql = 'select max(id) from {} where {} = "{}"'.format(table, flddupl, valdupl)
                cur.execute(sql)
                for row in cur:
                    id = row['max(id)']
                if id is None:
                    id = 0
            connection.commit()
        if int(id) == 0:  # dublia ne nashli ili i ne iskali tk poh
            if kwargs is not None:
                for key, value in kwargs.items():
                    fields = fields + key + ", "
                    vals = vals + "\"" + value + "\", "
            fields = fields[0:-2]
            vals = vals[0:-2]
            with connection.cursor() as cursor:
                sql = "INSERT INTO " + table + " (" + fields + ") VALUES (" + vals + ")"
                if debug == True:
                    print("ins : " + sql)
                cursor.execute(sql)
            connection.commit()
            with connection.cursor() as cur:
                sql = "select max(id) from " + table
                cur.execute(sql)
                for row in cur:
                    id = row['max(id)']
            connection.commit()
            if debug == True : print("ins id: " + str(id))
    except Exception as ex:
        id = -1
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)
    finally:
        return id


def upd(connection, table, id, **kwargs):
    try:
        vals = ""
        if kwargs is not None:
            for key, value in kwargs.items():
                vals = vals + key + " = \"" + str(value) + "\", "
        vals = vals[0:-2]
        with connection.cursor() as cursor:
            sql = "update " + table + " set " + vals + " where id = " + str(id)
            if debug == True:
                print("upd : " + sql)
            cursor.execute(sql)
        connection.commit()
    finally:
        sql = ""


def delete(connection, table, all=False, id=0):
    with connection.cursor() as cursor:
        sql = "delete FROM " + table
        if all != True:
            sql = sql + " WHERE id = " + str(id)
        if debug == True:
            print("del : " + sql)
        cursor.execute(sql)
        connection.commit()


def init_driver():
    driver = webdriver.Firefox()
    driver.wait = WebDriverWait(driver, WebDriverWaitInSec)
    return driver


def iselementpresent(bywhat, key):
    try:
        driver.find_element(bywhat, key)
        return True
    except:
        return False


def findelementtext(bywhat, key):
    try:
        txt = driver.find_element(bywhat, key).text
        return txt
    except:
        return ""


def findelementattr(bywhat, key, attr):
    try:
        txt = driver.find_element(bywhat, key).get_attribute(attr)
        return txt
    except:
        return ""


def lookup(driver, query, saveit):
    global num_, cat1_, cat2_, cat3_

    driver.get(query)
    try:
        driver.wait.until(EC.presence_of_element_located((By.CLASS_NAME, "company-map-mobile")))
        driver.implicitly_wait(WebDriverWaitInSec)

        name_ = clean(findelementtext(By.XPATH, "//h1[@class='title-company']"))
        descr_ = clean(findelementtext(By.XPATH, "//div[@class='description']"))
        geo_lat_ = clean(findelementattr(By.XPATH, "//meta[@itemprop='latitude']", "content"))
        geo_lon_ = clean(findelementattr(By.XPATH, "//meta[@itemprop='longitude']", "content"))
        site_ = clean(findelementtext(By.XPATH, "//div[@class='url']/a[@class='seo-link']/i"))
        email_ = clean(findelementtext(By.XPATH, "//div[@class='email']/a[@itemprop='email']/i"))
        phone_ = clean(findelementtext(By.XPATH, "//div[@class='main-phone']"))
        addr_ = clean(findelementtext(By.XPATH, "//span[@itemprop='streetAddress']"))
        region_ = clean(findelementtext(By.XPATH, "//span[@itemprop='addressLocality']"))
        ind_ = clean(findelementtext(By.XPATH, "//span[@itemprop='addressRegion']"))
        sched_ = clean(findelementtext(By.XPATH, "//div[@class='period']"))

        if saveit:
            ins(connection_, "data1", True, "sourcelink", query,  name=name_, descr=descr_, geo_lat=geo_lat_, geo_lon=geo_lon_, site=site_, email=email_, phone=phone_, addr=addr_, region=region_, ind=ind_, sched=sched_, num=str(num_), sourcelink=query, cat1=cat1_, cat2=cat2_, cat3=cat3_)

    except Exception as ex:
        template = "An exception of type {0} occurred. Arguments:\n{1!r}"
        message = template.format(type(ex).__name__, ex.args)
        print(message)


connection_ = getconnect()

if __name__ == "__main__":
    dcat1 = init_driver()
    dcat2 = init_driver()
    dcat3 = init_driver()
    dcat4 = init_driver()
    driver = init_driver()
    qlist = [
        "http://www.orgpage.ru/search.html?q=&loc=%D0%92%D0%B8%D0%B4%D0%BD%D0%BE%D0%B5&forReplies=false",
    ]
    cat1_ = ""
    cat2_ = ""
    cat3_ = ""
    for query in qlist:
        dcat1.get(query)
        listcat1 = dcat1.find_elements_by_xpath("//div[@class='row']/div[@class='col-lg-3 col-md-3 col-sm-6']/ul/li/a")
        for lcat1 in listcat1:
            cat1_ = lcat1.text
            lookup(dcat2, lcat1.get_attribute('href'), False)
            listcat2 = dcat2.find_elements_by_xpath("//div[contains(@class, 'col-lg-6 col-md-6 col-sm-6')]/ul/li/a")
            for lcat2 in listcat2:
                cat2_ = lcat2.text
                lookup(dcat3, lcat2.get_attribute('href'), False)
                listcat3 = dcat3.find_elements_by_xpath("//div[contains(@class, 'col-lg-6 col-md-6 col-sm-6')]/ul/li/a")
                for lcat3 in listcat3:
                    cat3_ = lcat3.text
                    lookup(dcat4, lcat3.get_attribute('href'), False)
                    list_of_links = dcat4.find_elements_by_xpath("//div[@class='wrp item-object clear-fix']/a")
                    for link in list_of_links:
                        lookup(driver, link.get_attribute('href'), True)
                        time.sleep(1)
                    time.sleep(1)
                time.sleep(1)
            time.sleep(1)
    dcat1.quit()
    dcat2.quit()
    dcat3.quit()
    dcat4.quit()
    driver.quit()

disconnect(connection_)


