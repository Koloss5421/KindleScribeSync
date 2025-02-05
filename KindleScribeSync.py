#!/bin/python3
#
# Kindle Scribe Sync
# Copyright (c) 2025 Koloss5421
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

import os, io, sys, json, time, pickle, pystray, logging, tarfile, img2pdf, schedule, requests, selenium.webdriver, selenium.webdriver.common, selenium.webdriver.firefox, selenium.webdriver.firefox.firefox_binary
from PIL import Image
from shutil import rmtree
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

## CONSTANTS
RENDER_HEIGHT = 2500
RENDER_WIDTH = 1200
NOTEBOOK_JSON_PATH = "notebooks.json"
COOKIES_FILE = "cookies.pkl"
UPDATE_MINUTES = 30

## Where to extract tar images
EXTRACT_PATH = "extraction"
## Where do we want to save the kindle notebooks?
SYNC_PATH = "kindle_notebooks"
FIREFOX_PATH = "C:\\Program Files\\Mozilla Firefox\\firefox.exe"
## User agent must be an android user agent to allow you to access the kindle-notebook site and see the api/notes page.
USER_AGENT = "Mozilla/5.0 (Linux; Android 11; SAMSUNG SM-G973U) AppleWebKit/537.36 (KHTML, like Gecko) SamsungBrowser/14.2 Chrome/87.0.4280.141 Mobile Safari/537.36"

AMZ_RENDER_HEADER = "x-amzn-karamel-notebook-rendering-token"
URL_AUTH = "https://read.amazon.com/kindle-notebook?ref_=neo_mm_yn_na_kfa"
URL_GET_NOTEBOOKS = "https://read.amazon.com/kindle-notebook/api/notes"
URL_OPEN_NOTEBOOK = "https://read.amazon.com/openNotebook?notebookId=[NOTEBOOK_ID]&marketplaceId=ATVPDKIKX0DER"
URL_RENDER_NOTEBOOK = "https://read.amazon.com/renderPage?startPage=0&endPage=[NOTEBOOK_LENGTH]&width={}&height={}&dpi=160".format(RENDER_WIDTH, RENDER_HEIGHT)

ff_options = selenium.webdriver.FirefoxOptions()
ff_profile = selenium.webdriver.FirefoxProfile()

## Selenium basic options
ff_options.binary_location = FIREFOX_PATH
ff_options.profile = ff_profile
ff_profile.set_preference("general.useragent.override", USER_AGENT)
ff_profile.set_preference("network.proxy.type", 2)
driver = selenium.webdriver.Firefox(ff_options)

## Setup Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(threadName)-12.12s] [%(levelname)-5.5s]  %(message)s",
    handlers=[
        logging.FileHandler("debug.log"),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger()

## GLOBALS
notebooks = {}
cookies = None
session = None
running = True
update_count = 0
last_update = "No Updates"

def load_notebook_json():
    """
    Loads the json file into the `notebooks` global
    """
    global notebooks
    logger.info("Attempting to load notebook data file")
    if os.path.exists(NOTEBOOK_JSON_PATH):
        logger.info("Loading notebook data file")
        with open(NOTEBOOK_JSON_PATH, "rb") as f:
            notebooks = json.load(f)

def save_notebook_json():
    """
    Saves the `notebooks` global to a file
    """
    global notebooks
    logger.info("Saving notebook data file")
    with open(NOTEBOOK_JSON_PATH, "w") as f:
        json.dump(notebooks, f)

def close_app():
    """
    shutdown the application
    """
    global driver
    global tray_icon
    global running
    running = False
    logger.info("Closing application")
    save_cookies()
    tray_icon.stop()
    try:
        sys.exit()
    except:
        pass

def update_info(icon, item):
    """
    callback for `pystray` menu item. `Last Update` button notification stating the last update time and number of items updated.
    """
    global last_update
    global update_count
    notify_string = "Last Updated: {} | Updated '{}' item(s)".format(last_update, str(update_count))
    icon.notify(notify_string)

def convert_to_pdf(images, savepath):
    """
    Uses `img2pdf` to convert an array of image paths (`images`) to a pdf (`savepath`).
    """
    logger.info("Converting extracted images to pdf output: {}".format(savepath))
    with open(savepath, "wb") as f:
        pdfdata = img2pdf.convert(images)
        f.write(pdfdata)

def extract_tarfile(tar_file_data):
    """
    Uses `tarfile` to extract the images from the amazon tar file. Returns an array of image paths.
    """
    logger.info("Extracting notebook tar file data")
    tar_stream = io.BytesIO(tar_file_data)
    tar_file = tarfile.open(fileobj=tar_stream)
    images = []
    for member in tar_file:
        if member.name.endswith(".png"):
            extr_path = "{}\\{}".format(EXTRACT_PATH, member.name)
            tar_file.extract(member, path=EXTRACT_PATH)
            images.append(extr_path)

    return images

def render_notebook(renderingToken, notebook_len):
    """
    Uses requests to call the render notebook url with the renderingToken. 
    The Response content should contain a tar file of the notebook.
    """
    global cookies
    global session
    logger.info("Rendering notebook")
    request_url = URL_RENDER_NOTEBOOK.replace("[NOTEBOOK_LENGTH]", str(notebook_len))
    session.headers[AMZ_RENDER_HEADER] = renderingToken
    while True:
        resp = session.get(request_url)
        if resp.is_redirect:
            rm_cookies()
            authenticate()
        else:
            break
    session.headers.pop(AMZ_RENDER_HEADER)
    cookies = requests.utils.dict_from_cookiejar(session.cookies)
    return resp.content

def get_notebook(id):
    """
    Uses requests to call the open notebook url to get the notebooks metadata and renderingToken.
    """
    global cookies
    global session
    logger.info("Getting notebook '{}' data".format(id))
    request_url = URL_OPEN_NOTEBOOK.replace("[NOTEBOOK_ID]", id)
    while True:
        resp = session.get(request_url)
        if resp.is_redirect:
            rm_cookies()
            authenticate()
        elif resp.status_code != 200:
            pass
        else:
            break
    cookies = requests.utils.dict_from_cookiejar(session.cookies)
    return resp.json()

def iterate_notebooks(obj, parentObj):
    """
    Recursive function that iterates over an items object, checking if the item is in a parent object.
    If the item is a folder, it creates the path and recurses.
    If the item is a notebook, it checks the updateTime against the modificationTime, if it is greater
    it updates the local copy of the notebook.
    """
    if 'items' in parentObj:
        parentItems = parentObj['items']
    else:
        parentItems = parentObj

    for x in obj:
        id = x['id']

        if not id in parentItems:
            if 'path' in parentObj:
                newPath = "{}\\{}".format(parentObj['path'], x['title'])
            else:
                newPath = "{}".format(x['title'])

            parentItems[id] = {
                'type': x['type'],
                'name': x['title'],
                'path': newPath,
                'updateTime': 0,
                'items': {}
            }
        
        if x['type'] == "folder":
            folder_path = "{}\\{}".format(SYNC_PATH, parentItems[id]['path'])
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
                os.removedirs
            iterate_notebooks(x['items'], parentItems[id])

        
        if x['type'] == "notebook":
            nb_data = get_notebook(id)
            time.sleep(1)
            if nb_data['metadata']['modificationTime'] > parentItems[id]['updateTime']:
                total_pages = nb_data['metadata']['totalPages']

                if (total_pages > 0):
                    total_pages = total_pages - 1

                tardata = render_notebook(nb_data['renderingToken'], total_pages)
                images = extract_tarfile(tardata)
                
                pdf_path = "{}\\{}.pdf".format(SYNC_PATH, parentItems[id]['path'])
                
                convert_to_pdf(images, pdf_path)

                for x in images:
                    os.remove(x)
                
                global update_count
                update_count += 1
                parentItems[id]['updateTime'] = int(time.time())

def id_exists_in_object(id, sync_items = []):
    """
    Checks if an id exists in an array object. Returns the index of the object.
    """
    for x in sync_items:
        if x["id"] == id:
            return sync_items.index(x)
    return -1


def prune_orphans(items, sync_items):
    """
    Recursive function that iterates over a set of items, if the item is not in the sync_items array it removes it based on the item type.
    If the item exists in both and is of type folder, it recurses on the next set of items.
    """
    global update_count
    for k in list(items.keys()):
        dict_object = items[k]
        index = id_exists_in_object(k, sync_items)
        if  index == -1:
            if dict_object["type"] == "folder":
                folder_path = "{}\\{}".format(SYNC_PATH, dict_object['path'])
                try:
                    logger.info("Pruning '{}' Folder".format(folder_path))
                    rmtree(folder_path)
                    update_count += 1
                except:
                    logger.error("Pruning '{}' Folder Failed!".format(folder_path))
            if dict_object["type"] == "notebook":
                pdf_path = "{}\\{}.pdf".format(SYNC_PATH, dict_object['path'])
                try:
                    logger.info("Pruning '{}' Notebook".format(pdf_path))
                    os.remove(pdf_path)
                    update_count += 1
                except:
                    logger.error("Pruning '{}' Notebook Failed!".format(pdf_path))
            del items[k]
        else:
            if dict_object["type"] == "folder":   
                prune_orphans(dict_object["items"], sync_items[index]["items"])

def get_all_notebooks():
    """
    Uses requests to query the notes api. The object returned contains the entire structure
    of notebooks from the scribe.
    """
    logger.info("Getting all notebooks")
    global cookies
    global session
    while True:
        resp = session.get(URL_GET_NOTEBOOKS)
        if resp.is_redirect:
            rm_cookies()
            authenticate()
        else:
            break
    cookies = requests.utils.dict_from_cookiejar(session.cookies)
    data = resp.json()
    iterate_notebooks(data['itemsList'], notebooks)
    prune_orphans(notebooks, data['itemsList'])
    save_notebook_json()

def load_cookies():
    """
    Uses `pickle` to load cookies from disk as the `cookies` global and update the requests session.
    """
    global cookies
    global session
    logger.info("Attempting to load cookies")
    if os.path.exists(COOKIES_FILE):
        logger.info("Loading Cookies from file")
        with open(COOKIES_FILE, "rb") as f:
            cookies = pickle.load(f)

        if cookies == None or len(cookies) < 5:
            return False
        
        session.cookies.update(cookies)
        return True
    
    return False

def save_cookies():
    """
    Uses `pickle` to save the global `cookies` object to disk for later use.
    """
    global cookies
    logger.info("Saving cookies")
    with open(COOKIES_FILE, "wb") as f:
        pickle.dump(cookies, f)

def rm_cookies():
    """
    Clears all cookie instances including the on disk file.
    """
    global session
    global cookies
    global driver
    logger.info("Deleting all cookies")
    driver.delete_all_cookies()
    session.cookies.clear()
    cookies = None
    if (os.path.exists(COOKIES_FILE)):
        os.remove(COOKIES_FILE)

def authenticate():
    """
    Uses selenium driver to allow you to authenticate to kindle, saves the cookies 
    for later use and updates the requests session.
    """
    global driver
    global cookies
    global session
    logger.info("Authenticating to {}".format(URL_AUTH))
    driver.get(URL_AUTH)

    try:
        logger.info("Waiting for authentication")
        element = WebDriverWait(driver, 120).until(
            EC.presence_of_element_located(
                (By.ID, "web-library-root")
            )
        )
    finally:
        pass

    cookies = driver.get_cookies()
    save_cookies()

    driver.quit()

    session.cookies.update(cookies)

def check_notebooks():
    """
    Uses as a scheduled job with schedule and the kickoff for all notebook syncs 
    including forced updates with the system tray icon.
    """
    global session
    global update_count
    global last_update
    update_count = 0
    logger.info("Checking for notebook changes")

    if session == None:
        session = requests.session()
        session.headers.update({'User-Agent': USER_AGENT})

    if cookies == None:
        if not load_cookies():
            authenticate()
        else:
            driver.quit()

    get_all_notebooks()

    now = datetime.now()
    last_update = now.strftime("%m/%d/%Y, %H:%M:%S")

    logger.info("Will Check again in {} minutes...".format(UPDATE_MINUTES))


if not os.path.exists(EXTRACT_PATH):
    os.mkdir(EXTRACT_PATH)

## Setup the system tray icon / menus
menu = pystray.Menu(
    pystray.MenuItem('Last Update', update_info),
    pystray.MenuItem('Force Sync', check_notebooks),
    pystray.MenuItem('Quit', close_app)
)

tray_image = Image.open("KindleScribeSyncIcon.png")
tray_icon = pystray.Icon("Kindle Scribe Sync", tray_image, "Kindle Scribe Sync", menu)  

## Load the notebook json object
load_notebook_json()

logger.info("Running Tray Icon")
tray_icon.run_detached()

logger.info("Running initial check")
check_notebooks()

## schedule the check_notebooks function
schedule.every(UPDATE_MINUTES).minutes.do(check_notebooks)

## Run forever please.
while running:
    schedule.run_pending()
    time.sleep(1)