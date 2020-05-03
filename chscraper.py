import json
import os
import urllib
import xlsxwriter
import argparse
import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from threading import Thread

def find_element_by_id(source, id):
    ''' 
    Find first matched element by ID from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @id The ID of web element to search
    '''
    if source is None: return None
    try:
        return source.find_element_by_id(id)
    except:
        return None

def find_element_by_tag_name(source, name):
    ''' 
    Find first matched element by tag name from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @name The tag name of web element to search
    '''
    try:
        return source.find_element_by_tag_name(name)
    except:
        return None

def find_element_by_class_name(source, cname):
    ''' 
    Find first matched element by class name from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @cname The class name of web element to search
    '''
    if source is None: return None
    try:
        return source.find_element_by_class_name(cname)
    except:
        return None

def find_element_by_css_selector(source, selector):
    ''' 
    Find first matched element by CSS selector from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @selector The CSS selector string to search the web element
    '''
    if source is None: return None
    try:
        return source.find_element_by_css_selector(selector)
    except:
        return None

def find_element_by_name(source, name):
    ''' 
    Find first matched element by name selector from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @name The name of web element to search.
    '''
    if source is None: return None
    try:
        return source.find_element_by_name(name)
    except:
        return None

def find_element_by_link_text(source, text, partial = True):
    ''' 
    Find first matched element by hyperlink text from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @text The text in the hyperlink element to search
    @partial Set True to find by partial text or false to find by exact text
    '''
    if source is None: return None
    try:
        if partial:
            return source.find_element_by_partial_link_text(text)
        else:
            return source.find_element_by_link_text(text)
    except:
        return None

def find_element_by_xpath(source, xpath):
    ''' 
    Find first matched element by XPath selector from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @name The name of web element to search.
    '''
    if source is None: return None
    try:
        return source.find_element_by_xpath(xpath)
    except:
        return None

def find_elements_by_tag_name(source, name):
    ''' 
    Find all matched element by tag name from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @name The tag name of web element to search
    '''
    if source is None: return None
    try:
        return source.find_elements_by_tag_name(name)
    except:
        return None

def find_elements_by_class_name(source, cname):
    ''' 
    Find all matched element by class name from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @cname The class name of web element to search
    '''
    if source is None: return None
    try:
        return source.find_elements_by_class_name(cname)
    except:
        return None

def find_elements_by_css_selector(source, selector):
    ''' 
    Find all matched element by CSS selector from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @selector The CSS selector string to search the web element
    '''
    if source is None: return None
    try:
        return source.find_elements_by_css_selector(selector)
    except:
        return None

def find_elements_by_name(source, name):
    ''' 
    Find all matched element by name selector from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @name The name of web element to search.
    '''
    if source is None: return None
    try:
        return source.find_elements_by_name(name)
    except:
        return None

def find_elements_by_link_text(source, text, partial = True):
    ''' 
    Find all matched element by hyperlink text from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @text The text in the hyperlink element to search
    @partial Set True to find by partial text or false to find by exact text
    '''
    if source is None: return None
    try:
        if partial:
            return source.find_elements_by_partial_link_text(text)
        else:
            return source.find_elements_by_link_text(text)
    except:
        return None

def find_elements_by_xpath(source, xpath):
    ''' 
    Find all matched element by XPath selector from driver or web element and returns None if not found.
    @source Either webdriver engine or web element
    @name The name of web element to search.
    '''
    if source is None: return None
    try:
        return source.find_elements_by_xpath(xpath)
    except:
        return None

def minval(left, right):
    '''
    Find the minimum value between integers or float number
    '''
    if left < right: return left
    return right

def maxval(left, right):
    '''
    Find the maximum value between integers or float number
    '''
    if left > right: return left
    return right

def isundefined(value):
    '''
    Determines whether the given value is None, empty string or empty list
    '''
    if value is None:
        return True
    if isinstance(value, str):        
        return len(value) < 1
    if isinstance(value, list):
        return len(value) < 1
    return False

def convertlist(data):
    '''
    Try to convert the given data into a list data type, returns empty list if failed.
    '''
    if isinstance(data, list):
        return data
    elif isinstance(data, str):
        parts = data.split(';')
        array = []
        for node in parts:
            if not node is None and len(node) > 0:
                node = node.trim()
                if len(node) > 0:
                    array.append(node)
        return array
    else:
        return []

def convertstr(data):
    '''
    Convert the given data to a string and returns empty string if failed.
    '''
    if data is None:
        return ''
    elif isinstance(data, str):
        return data
    elif isinstance(data, list):
        result = ''
        for node in data:
            if not node is None:
                if isinstance(node, str):
                    result += node + ";"
                elif isinstance(node, int) or isinstance(node, bool) or isinstance(node, float):
                    result += str(node) + ";"
                elif isinstance(node, list):
                    result += convertstr(node) + ";"
                else:
                    result += str(node) + ";"
        result = result.strip(';')
        return result
    else:
        return str(data)

def convertbool(data):
    '''
    Convert the given data to a boolean value and returns False when failed
    '''
    if data is None:
        return False
    elif isinstance(data, bool):
        return data
    elif isinstance(data, str):
        value = data.lower()
        if value == 'true' or value == '1' or value == 'ok' or value == 'on': return True
        return False
    elif isinstance(data, int):
        return data != 0
    elif isinstance(data, float):
        return data != 0
    else:
        return False

def convertint(data):
    '''
    Convert the given data to an integer and returns zero when failed
    '''
    if data is None:
        return 0
    elif isinstance(data, int):
        return data
    elif isinstance(data, float):
        return int(data)
    elif isinstance(data, str):
        return int(data)
    elif isinstance(data, bool):
        return 1 if data != False else 0
    else:
        return 0

def parse_args(args = None):
    '''
    Parse the given command arguments list into the dictionary objects
    '''
    if args is None or isundefined(args):
        return None
    elif isinstance(args, dict):
        return args
    elif isinstance(args, list):
        parser = argparse.ArgumentParser(description="Companies house site web scraper configuration arguments")
        parser.add_argument("--query", type = str, help = "The company names to search, use comma as separator of companies.", required = True, metavar = "string")
        parser.add_argument("--output", type = str, help = "The output directory where the scraping results will saved.", required = True, metavar = "path")
        parser.add_argument("--limit", type = int, help = "Optional, the maximum number of records to scrap, set with zero or omit this argument to scrap all records. Default to 10 records.", default = 10, required=True, metavar = "number")
        parser.add_argument("--pages", type = int, help = "Optional, the maximum number of pages to scrap, set with zero to scrap all pages. Default to 1 page.", default = 1, required = False, metavar = "number")
        parser.add_argument("--hidden", type = bool, help='Optional, set with True to make headless (invisible) scraper browser, otherwise, set False to show the scraper browser (debug mode).', default = True, metavar= 'boolean')
        parser.add_argument("--threads", type = int, help = "Optional, the maximum number of threads to be used by scraper.", default = os.cpu_count(), required = False, metavar = "number")
        parser.add_argument("--histories", type = bool, help = "Optional, set True (default) to scrap company histories or set False to not scrap it.", default = True, required = False , metavar = "boolean")
        parser.add_argument("--officers", type = bool, help = "Optional, set True (default) to scrap company officers or set False to not scrap it.", default = True, required = False, metavar = "boolean")
        parser.add_argument("--binary", type = str, help = "The path to the driver executable binary file to be used, omit this parameter to use default path.", required = False, metavar ="path")
        parser.add_argument("--options", type = str, help = "The driver arguments list to use, use comma as separator between arguments.", required = False, metavar = "string")
        parser.add_argument("--exclusion", type = str, help =" The driver exclusion argument list to use, use comma as separator between arguments.", required = False, metavar = "string")
        return vars(parser.parse_args(args))
    
    else:
        return None

def runtime_config():
    '''
    Read and parse the active runtime settings dictionary for the current instance.
    If the command line arguments is supplied, then it parsed and used as settings.
    Otherwise, this function will search and load the persistent configuration file.
    '''
    cmd = os.sys.argv
    if len(cmd) > 1:
        cmd.remove(cmd[0])
        map = parse_args(cmd)
        if not map is None:
            result = dict()
            result['driver_options'] = None
            result['driver_exclude'] = None
            result['driver_extensions'] = None
            result['driver_binary'] = None
            result['driver_logpath'] = None
            result['driver_hidden'] = True
            result['driver_address'] = None
            result['company_names'] = None
            text = map['query']
            if not isundefined(text):
                temp = str(text).split(",")
                data = []
                for sub in temp:
                    if not isundefined(sub):
                        data.append(sub.strip())
                result['company_names'] = data
            if isundefined(result['company_names']):
                raise Exception('The "--query" argument is needed and must contains one or more company names to query, every of queries must separated using commas.')
            text = convertstr(map['output'])
            if isundefined(text):
                raise Exception('The "--output" argument is needed and must define either relative or absolute path to the output directory where the scraped results will saved to.')
            result['output_folder'] = text
            text = map['options']
            if not isundefined(text):
                temp = str(text).split(",")
                data = []
                for sub in temp:
                    if not isundefined(sub):
                        data.append(sub.strip())
                result['driver_options'] = data
            if isundefined(result['driver_options']):
                result['driver_options'] = ["--disable-blink-features", "--disable-blink-features=AutomationControlled"]
            text = map['exclusion']
            if not isundefined(text):
                temp = str(text).split(",")
                data = []
                for sub in temp:
                    if not isundefined(sub):
                        data.append(sub.strip())
                result['driver_exclude'] = data
            if isundefined(result['driver_exclude']):
                result['driver_exclude'] = ['enable-automation', 'enable-logging']
            result['driver_extensions'] = None
            result['driver_hidden'] = convertbool(map['hidden'])
            result['driver_binary'] = convertstr(map['binary'])
            result['driver_address'] = ''
            result['driver_logpath'] = ''
            result['scrap_website'] = 'https://beta.companieshouse.gov.uk'
            num = map['limit']
            if isundefined(num) or not isinstance(num, int) or num < 1:
                result['scrap_limits'] = 0
            else:
                result['scrap_limits'] = num
            num = map['pages']
            if isundefined(num) or not isinstance(num, int) or num < 1:
                result['maximum_pages'] = 0
            else:
                result['maximum_pages'] = num
            result['scrap_logging'] = True
            num = map['threads']
            if isundefined(num) or not isinstance(num, int) or num < 1:
                result['scrap_parallel'] = 0
            else:
                result['scrap_parallel'] = num
            result['crawl_overview'] = True
            state = map['histories']
            if isundefined(state):
                state = True
            elif isinstance(state, int):
                if state != 0: state = True
                else: state = False
            elif isinstance(state, str):
                if state.lower() == 'false': state = False
                else: state = True
            elif not isinstance(state, bool):
                state = True
            result['crawl_histories'] = convertbool(state)
            state = map['officers']
            if isundefined(state):
                state = True
            elif isinstance(state, int):
                if state != 0: state = True
                else: state = False
            elif isinstance(state, str):
                if state.lower() == 'false': state = False
                else: state = True
            elif not isinstance(state, bool):
                state = True
            result['crawl_officers'] = convertbool(state)
            return result
    paths = [os.path.abspath("settings.json"), os.path.abspath("setting.json"), os.path.abspath("config.json"), os.path.abspath("scrap.json"), os.path.abspath("scraper.json"), os.path.abspath("chscraper.json")]
    for path in paths:
        if os.path.exists(path):
            try:
                with open(path, 'r') as file:
                    data = file.read()
                    if len(data) > 0:
                        map = json.loads(data)
                        if not map is None and isinstance(map, dict):
                            return map
            except:
                continue
    return []

def get_company_code(li, a):
    href = a.get_attribute("href")
    if not isundefined(href):
        exploded = href.split('/')
        if len(exploded) > 2 and exploded[len(exploded) - 2] == "company":
            return exploded[len(exploded) - 1]
    p = find_element_by_tag_name(li, 'p')
    if not p is None:
        strong = find_element_by_tag_name(p, 'strong')
        if not strong is None:
            text = strong.text
            if not text is None and len(text) > 0:
                return text
    return None

def get_percent_int(pos: int, len: int):
    fpos = float(pos)
    flen = float(len)
    fprc = (fpos / flen) * 100
    return int(round(fprc, 0))

def get_percent_flo(pos: int, len: int):
    fpos = float(pos)
    flen = float(len)
    fprc = (fpos / flen) * 100
    return round(fprc, 2)

def wait_until_end(workers):
    while True:
        done = False
        for task in workers:
            done = task.is_ended and task.has_result
            if done: break
        if not done:
            time.sleep(float(0.25))
        if done: break

def join_list(sequence, separator):
    result = ''
    index = 0
    count = len(sequence)
    for node in sequence:
        result += str(node)
        if index + 1 < count:
            result += separator
        index += 1
    return result

class ThreadTask ( Thread ):
    def __init__(self, target, query, node, tname):
        Thread.__init__(self, target = target, name = "Thread-" + str(tname))
        self._tname = tname
        self._wname = query
        self._node = node
        self._return = None
        self._is_done = False

    def is_ended(self):
        return self._is_done == True and not self.isalive()

    def has_result(self):
        return not self._return is None

    def get_result(self):
        return self._return

    def run(self):
        if self._target is not None:
            self._return = self._target(self._wname, self._node, "Thread-" + str(self._tname))
            self._is_done = True

    def join(self, *args):
        Thread.join(self, *args)
        return self._return

class ScrapSettings:
    
    def __init__(self, queries = None, output = None):
        self.dvargs = ["--disable-blink-features", "--disable-blink-features=AutomationControlled"]
        self.dvexcl = ["enable-automation", "enable-logging"]
        self.dvexts = []
        self.dvhide = True
        self.dvexec = ''
        self.dvaddr = ''
        self.dvlogs = ''
        self.landing = 'https://beta.companieshouse.gov.uk'
        self.mrows = 20
        self.mpage = 5
        self.thread = os.cpu_count()
        self.logging = True
        self.history = True
        self.officer = True
        self.queries = convertlist(queries)
        self.exactly = True
        text = convertstr(output)
        if not isundefined(text):
            if not os.path.isabs(text):
                self.output = os.path.abspath(text)
            else:
                self.output = text
        else:
            self.output = os.path.abspath('output')
        self.useapi = False
        self.apikey = ''

    def reload(self):
        data = runtime_config()
        self.dvargs = convertlist(data.get('driver_options', None))
        self.dvexcl = convertlist(data.get('driver_exclude', None))
        self.dvexts = convertlist(data.get('driver_extensions', None))
        self.dvhide = convertbool(data.get('driver_hidden', True))
        self.dvexec = convertstr(data.get('driver_binary', None))
        self.dvaddr = convertstr(data.get('driver_address', None))
        self.dvlogs = convertstr(data.get('driver_logpath', None))
        self.queries = convertlist(data.get('company_names', None))
        self.landing = convertstr(data.get('scrap_website', 'https://beta.companieshouse.gov.uk'))
        self.mrows = convertint(data.get('scrap_limits', 0))
        self.mpage = convertint(data.get('maximum_pages', 0))
        self.thread = convertint(data.get('scrap_parallel', os.cpu_count()))
        self.logging = convertbool(data.get('scrap_logging', True))
        self.output = convertstr(data.get('output_folder', True))
        self.history = convertbool(data.get('crawl_histories', True))
        self.officer = convertbool(data.get('crawl_officers', True))
        self.exactly = convertbool(data.get('exact_matches', True))
        self.useapi = convertbool(data.get('restapi_enable', None))
        self.apikey = convertstr(data.get('restapi_token', None))
        

    def exports(self):
        map = dict()
        map['driver_options'] = self.dvargs
        map['driver_exclude'] = self.dvexcl
        map['driver_extensions'] = self.dvexts
        map['driver_hidden'] = self.dvhide
        map['driver_binary'] = self.dvexec
        map['driver_address'] = self.dvaddr
        map['driver_logpath'] = self.dvlogs

        map['company_names'] = self.queries
        map['scrap_website'] = self.landing
        map['scrap_limits'] = self.mrows
        map['scrap_parallel'] = self.thread
        map['scrap_logging'] = self.logging
        map['exact_matches'] = self.exactly
        map['maximum_pages'] = self.mpage
        map['output_folder'] = self.output
        map['crawl_histories'] = self.history
        map['crawl_officers'] = self.officer
        map['restapi_enable'] = self.useapi
        map['restapi_token'] = self.apikey
        return map

    def serialize(self):

        return json.dumps(self.export())

    def defaults(self, force = False):
        self.dvargs = ["--disable-blink-features", "--disable-blink-features=AutomationControlled"]
        self.dvexcl = ["enable-automation", "enable-logging"]
        self.dvexts = []
        self.dvhide = True
        self.dvexec = ''
        self.dvaddr = ''
        self.dvlogs = ''
        self.landing = 'https://beta.companieshouse.gov.uk'
        self.mrows = 20
        self.mpage = 5
        self.thread = os.cpu_count()
        self.logging = True
        self.history = True
        self.officer = True
        self.exactly = True
        if force:
            self.queries = []
            self.output = os.path.abspath('output')
        self.useapi = False
        self.apikey = ''

    def prepare(self):
        if isundefined(self.output):
            raise Exception('The output folder is not defined, please define it first either from "--output" parameter in command line arguments or "output_folder" in configuration file.')
        if isundefined(self.queries):
            raise Exception('The company names is not defined, please define it first either from "--query" parameter in command line arguments or "company_names" in configuration file.')
        if isundefined(self.dvargs):
            self.dvargs = ["--disable-blink-features", "--disable-blink-features=AutomationControlled"]
        if isundefined(self.dvexcl):
            self.dvexcl = ["enable-automation", "enable-logging"]
        if isundefined(self.landing):
            self.dvexec = 'https://beta.companieshouse.gov.uk'
        if isundefined(self.mrows) or self.mrows < 0:
            self.mrows = 20
        if isundefined(self.mpage) or self.mpage < 0:
            self.mpage = 0
        if isundefined(self.thread) or self.thread < 1:
            self.thread = os.cpu_count()
        return self

    def cfgload(self, path):
        if isundefined(path):
            raise Exception('The "path" is not defined, make sure the "path" is represent an absolute or relative path to the configuration file.')
        if not isinstance(path, str):
            path = str(path)
        if not os.path.isabs(path):
            path = os.path.abspath(path)
        if not os.path.exists(path):
            raise Exception('The configuration file is not found at following path: "' + path + '".')
        if not os.path.isfile(path):
            raise Exception('The path is not represent a configuration file path, it seems refer to directory path: "' + path + '".')
        with open(path, 'r') as file:
            source = file.read()
            data = json.loads(source)
            self.dvargs = convertlist(data.get('driver_options', None))
            self.dvexcl = convertlist(data.get('driver_exclude', None))
            self.dvexts = convertlist(data.get('driver_extensions', None))
            self.dvhide = convertbool(data.get('driver_hidden', True))
            self.dvexec = convertstr(data.get('driver_binary', None))
            self.dvaddr = convertstr(data.get('driver_address', None))
            self.dvlogs = convertstr(data.get('driver_logpath', None))
            self.queries = convertlist(data.get('company_names', None))
            self.landing = convertstr(data.get('scrap_website', 'https://beta.companieshouse.gov.uk'))
            self.mrows = convertint(data.get('scrap_limits', 0))
            self.mpage = convertint(data.get('maximum_pages', 0))
            self.thread = convertint(data.get('scrap_parallel', os.cpu_count()))
            self.logging = convertbool(data.get('scrap_logging', True))
            self.output = convertstr(data.get('output_folder', True))
            self.history = convertbool(data.get('crawl_histories', True))
            self.officer = convertbool(data.get('crawl_officers', True))
            self.exactly = convertbool(data.get('exact_matches', False))
            self.useapi = convertbool(data.get('api_enable', None))
            self.apikey = convertstr(data.get('api_token', None))
        return True

    def cfgsave(self, path):
        if isundefined(path):
            raise Exception('The "path" is not defined, make sure the "path" is represent an absolute or relative path to the configuration file.')
        if not isinstance(path, str):
            path = str(path)
        if not os.path.isabs(path):
            path = os.path.abspath(path)
        if not os.path.isfile(path):
            raise Exception('The path is not represent a configuration file path, it seems refer to directory path: "' + path + '".')
        folder = os.path.dirname(path)
        if not os.path.isabs(folder):
            folder = os.path.abspath(folder)
        if not os.path.exists(folder):
            os.makedirs(folder, 0o777, True)
        with open(path, 'w') as file:
            file.write(self.serialize())
            file.flush()
        return path

    def options(self, chrome: bool = True):
        if chrome:
            opt = webdriver.ChromeOptions()
            if not isundefined(self.dvargs):
                for node in self.dvargs:
                    opt.add_argument(node)
            if not isundefined(self.dvexcl):
                opt.add_experimental_option("excludeSwitches", self.dvexcl)
            if not isundefined(self.dvexts):
                for ext in self.dvexts:
                    opt.add_extension(ext)
            if not isundefined(self.dvaddr):
                opt.debugger_address = self.dvaddr
            if not isundefined(self.dvexec):
                opt.binary_location = self.dvexec
            opt.headless = self.dvhide
            return opt
        else:
            opt = webdriver.FirefoxOptions()
            if not isundefined(self.dvargs):
                for node in self.dvargs:
                    opt.add_argument(opt)
            if not isundefined(self.dvexec):
                opt.binary_location = self.dvexec
            opt.headless = self.dvhide
            return opt
    
    def chrome(self):
        opt = self.options(True)
        if not isundefined(self.dvlogs):
            driver = webdriver.Chrome(options=opt, service_log_path=self.dvlogs)
        else:
            driver = webdriver.Chrome(options=opt)
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", { "source": """Object.defineProperty(navigator, 'webdriver', { get: () => undefined }) """ }) 
        driver.execute_cdp_cmd("Network.enable", {})
        driver.execute_cdp_cmd("Network.setExtraHTTPHeaders", {"headers": {"User-Agent": "browser1"}})
        return driver

    def firefox(self):

        return webdriver.Firefox(self.options(False))

class ScrapProvider:

    def __init__(self, config: ScrapSettings):
        if config is None:
            config = ScrapSettings()
            config.reload()
        config.prepare()
        self.config = config

    def __scrapPage(self, driver, query: str):
        first = time.time()
        driver.get(self.config.landing.rstrip("/") + "/search/companies?q=" + urllib.parse.quote(query))
        expect = 10000
        matches = -1
        smeta = find_element_by_id(driver, 'search-meta')
        if not smeta is None:
            para = find_element_by_tag_name(smeta, 'p')
            if not para is None:
                ptext = para.text.strip()
                ptext = ptext.replace('matches found', "").strip()
                ptext = ptext.replace(",", "")
                if not isundefined(ptext):
                    try:
                        matches = int(expect)
                        expect = 0 if matches == 0 else int(round(matches / 20, 0))
                    except: expect = 10000
        mrows = self.config.mrows
        mpage = self.config.mpage
        paging = expect if mpage < 1 else minval(mpage, 10000)
        if matches == -1:
            print('> We could not found count of companies that available for query "' + query + '", we will run force mode with max pages = ' + str(paging) + ".")
        elif matches == 0:
            print('> The search with query "' + query + '" did not yield any results.')
            return []
        else:
            print('> The search with query "' + query + '" has found ' + str(matches) + ' records with expected ' + str(paging) + " pages to scrap.")
        result = []
        count = 0
        npage = 0
        qlower = query.lower()
        for number in range(0, paging, 1):
            number += 1
            if number > 1:
                driver.get(self.config.landing.rstrip("/") + "/search/companies?q=" + urllib.parse.quote(query) + "&page=" + str(number))
            ecode = find_element_by_id(driver, 'error-code')
            if not ecode is None:
                print("--- [Page " + str(number - 1) + "] 100% completed (no more pages are available)..")
                break
            prog = get_percent_flo(number, paging)
            print("--- [Page " + str(number) + "] " + str(prog) + "% completed..")
            npage = number
            clusters = find_element_by_id(driver, 'results')
            if clusters is None:
                clusters = find_element_by_class_name(driver, 'results-list')
                if not clusters:
                    continue
            litags = find_elements_by_tag_name(clusters, "li")
            rows = 0
            for li in litags:
                anchor = find_element_by_tag_name(li, "a")
                if not anchor is None:
                    title = anchor.text.strip()
                    if not self.config.exactly and title.lower().find(qlower) == -1:
                        continue
                    code = get_company_code(li, anchor)
                    if code is None:
                        continue
                    if mrows > 0 and count + 1 > mrows:
                        print("> Maximum records has reached.")
                        break
                    node = dict()
                    node['page'] = number
                    node['rows'] = rows
                    node['index'] = count + 1
                    node['code'] = code
                    node['name'] = anchor.text.strip()
                    node['href'] = anchor.get_attribute("href")
                    print("------ [" + str(count).zfill(6) + " in page " + str(number).zfill(3) + " at row " + str(rows + 1).zfill(2) + "] " + title + " (" + code + ")")
                    result.append(node)
                    count += 1
                    rows += 1
            if rows == 0:
                print("> Search terminated, no more records can be founded.")
                break
        e = int(time.time() - first)
        print("> Finally, we've found " + str(len(result)) + ' records with query "' + query + '", elapsed time = ' + '{:02d}:{:02d}:{:02d}'.format(e // 3600, (e % 3600 // 60), e % 60))
        return result

    def __scrapUser(self, driver, target):
        code = target['code']
        driver.get("https://beta.companieshouse.gov.uk/company/" + code + "/officers")
        container = find_element_by_class_name(driver, "appointments-list")
        if container is None:
            return []
        sections = find_elements_by_tag_name(container, "div")
        if sections is None or len(sections) < 1:
            return []
        index = 1
        result = []
        for div in sections:
            cname = div.get_attribute("class")
            if not cname is None and cname.startswith("appointment"):
                oname = find_element_by_id(div, "officer-name-" + str(index))
                oaddr = find_element_by_id(div, "officer-address-value-" + str(index))
                orole = find_element_by_id(div, "officer-role-" + str(index))
                ostat = find_element_by_id(div, "officer-status-tag-" + str(index))
                obirth = find_element_by_id(div, "officer-date-of-birth-" + str(index))
                oapdate = find_element_by_id(div, "officer-appointed-on-" + str(index))
                onat = find_element_by_id(div, "officer-nationality-" + str(index))
                oresi = find_element_by_id(div, "officer-country-of-residence-" + str(index))
                oocu = find_element_by_id(div, "officer-occupation-" + str(index))
                count = 0
                data = {}
                if not oname is None:
                   data['name'] = oname.text.strip()
                   count += 1
                if not ostat is None:
                    data['status'] = ostat.text.strip()
                    count += 1
                if not oaddr is None:
                    data['address'] = oaddr.text.strip()
                    count += 1
                if not orole is None:
                    data['role'] = orole.text.strip()
                    count += 1
                if not obirth is None:
                    data['birth'] = obirth.text.strip()
                    count += 1
                if not onat is None:
                    data['nationality'] = onat.text.strip()
                    count += 1
                if not oresi is None:
                    data['residence'] = oresi.text.strip()
                    count += 1
                if not oocu is None:
                    data['occupation'] = oocu.text.strip()
                    count += 1
                if not oapdate is None:
                    if not ostat is None and ostat.text.strip().lower() == 'resigned':
                        data['resigned'] = oapdate.text.strip()
                    else:
                        data['appointed'] = oapdate.text.strip()
                    count += 1
                if count > 0:
                    result.append(data)
                index += 1
        return result

    def __scrapHist(self, driver, target):
        code = target['code']
        driver.get("https://beta.companieshouse.gov.uk/company/" + code + "/filing-history")
        container = find_element_by_id(driver, "filing-history-content")
        if container is None:
            return []
        table = find_element_by_id(container, "fhTable")
        if table is None:
            table = find_element_by_tag_name(container, "table")
            if table is None:
                return []
        output = []
        rows = find_elements_by_tag_name(table, "tr")
        index = 1
        for row in rows:
            th = find_element_by_tag_name(row, "th")
            if th is None:
                tdlist = find_elements_by_tag_name(row, "td")
                if not tdlist is None and len(tdlist) > 2:
                    data = { "no": index , "date": tdlist[0].text.strip() }
                    offset = 1
                    tdnext = tdlist[offset]
                    tdclass = tdnext.get_attribute("class")
                    if not tdclass is None and tdclass.find("js-hidden") != -1:
                        tdnext = tdlist[2]
                        offset = 2
                    data["desc"] = tdnext.text.strip()
                    tdnext = tdlist[offset + 1]
                    if not tdnext is None:
                        a = find_element_by_tag_name(tdnext, "a")
                        if not a is None:
                            data["docs"] = a.get_attribute("href")
                    index += 1
                    output.append(data)
        return output

    def __scrapView(self, driver, target):
        container = find_element_by_id(driver, "content-container")
        if container is None: 
            return {}
        array = {}
        cpstat = find_element_by_id(driver, "company-status")
        csdate = find_element_by_id(driver, "cessation-date")
        cptype = find_element_by_id(driver, "company-type")
        cpcrdt = find_element_by_id(driver, "company-creation-date")
        array['name'] = target['name']
        dlList = find_elements_by_tag_name(container, "dl")
        if not dlList is None:
            found = False
            for dl in dlList:
                if found:
                    break
                dt = dl.find_element_by_tag_name("dt")
                if not dt is None and dt.text.strip().lower().find("address") != -1:
                    dd = find_element_by_tag_name(dl, "dd")
                    if not dd is None:
                        array['address'] = dd.text.strip()
                        found = True
        if not cpstat is None:
            array['status'] = cpstat.text.strip()
        if not cptype is None:
            array['type'] = cptype.text.strip()
        if not csdate is None:
            array['dissolved'] = csdate.text.strip()
        if not cpcrdt is None:
            array['incorporated'] = cpcrdt.text.strip()
        return array

    def __scrapMain(self, driver, target, tname):
        code = target['code']
        driver.get(self.config.landing.strip("/") + "/company/" + code)
        enode = find_element_by_id(driver, 'page-not-found-header')
        if not enode is None: 
            print('> The company with code "' + code + '" is not found..')
            return None
        result = { 'number': target['index'], 'paging': target['page'], "company": target['name'], 'identity': target['code'], 'profile': target['href'] }
        print("> " + tname + " => Scraping company overview information (" + target['name'] + ").")
        result['overview'] = self.__scrapView(driver, target)
        if self.config.history:
            print("> " + tname + " => Scraping company history information (" + target['name'] + ").")
            result['histories'] = self.__scrapHist(driver, target)            
        if self.config.officer:
            print("> " + tname + " => Scraping company officers information (" + target['name'] + ").")
            result['officers'] = self.__scrapUser(driver, target)
        return result

    def __scrapTask(self, query, target, tname):
        print("> " + tname + " => Scraping information about \"" + target['name'] + "\".")
        driver = self.config.chrome()
        try:
            return self.__scrapMain(driver, target, tname)
        finally:
            driver.close()
            driver.quit()

    def __scrapNode(self, query: str):
        driver = self.config.chrome()
        targets = self.__scrapPage(driver, query)
        driver.close()
        driver.quit()
        array = dict()
        header = dict()
        header['keywords'] = query
        header['datetime'] = datetime.datetime.now().isoformat()
        header['companies'] = 0
        header['elapsed'] = '00:00:00'
        if len(targets) == 0:
            array['header'] = header
            return array
        mtask = self.config.thread
        if mtask < 1: mtask = os.cpu_count()
        results = []
        if mtask == 1:
            for target in targets:
                res = self.__scrapTask(query, target, "Thread-1")
                if not res is None:
                    results.append(res)
            array['header'] = header
            array['matches'] = results
            return array
        else:
            workers = []
            count = 0
            print("> Maximum thread queue is: " + str(mtask))
            for target in targets:
                avail = len(workers)
                print("> Available queue: " + str(avail))
                if avail >= mtask:
                    print("> Waiting queue, total: " + str(avail))
                    while True:
                        done = False
                        for task in workers:
                            done = task._is_done
                            if not done: break
                        if done: break
                    for task in workers:
                        res = task.get_result()
                        if not res is None:
                            results.append(res)
                            count += 1
                    workers.clear()
                thread = ThreadTask(self.__scrapTask, query, target, len(workers) + 1)
                workers.append(thread)
                thread.start()
            if len(workers) > 0:
                while True:
                    done = False
                    for task in workers:
                        done = task._is_done
                        if not done: break
                    if done: break
                for task in workers:
                    res = task.get_result()
                    if not res is None:
                        results.append(res)
                        count += 1
                    workers.clear()
            header['companies'] = len(results)
            array['header'] = header
            array['matches'] = results
            return array

    def dispatch(self):
        first = time.time()
        result = dict()
        header = dict()
        header['queries'] = self.config.queries
        found = []
        array = []
        index = 1
        for query in self.config.queries:
            data = self.__scrapNode(query)
            if not data is None and len(data) > 0:
                matches = data.get('matches', None)
                if not matches is None:
                    for node in matches:
                        node['number'] = index
                        index += 1
                    data['matches'] = matches
                    array.append(data)
                    found.append(len(matches))
                else:
                    found.append(0)
            else:
                found.append(0)
        e = int(time.time() - first)
        header['elapsed'] = '{:02d}:{:02d}:{:02d}'.format(e // 3600, (e % 3600 // 60), e % 60)
        header['matches'] = found
        result['reports'] = header
        result['results'] = array
        jpath = self.__writeJson(result)
        xpath = self.__writeExcel(result)
        return [jpath, xpath]

    def __writeJson(self, output):
        folder = self.config.output
        if not os.path.isabs(folder):
            folder = os.path.abspath(folder)
        path = os.path.join(folder, "results.json")
        with open(path,'w') as file:
            file.write(json.dumps(output, indent = 4))
        return path
    
    def __writeExcel(self, output):
        folder = self.config.output
        if not os.path.isabs(folder):
            folder = os.path.abspath(folder)
        path = os.path.join(folder, "results.xlsx")
        book = xlsxwriter.Workbook(path)
        fmtinfo = book.add_format({'align': 'left', "bold": True, "border": 1})
        fmttitl = book.add_format({'align': 'left', "bold": True, "border": 1, 'bg_color': '#DDDDDD'})
        fmtkey = book.add_format({'align': 'left', "bold": False, "border": 1})
        fmttbl = book.add_format({'align': 'center', "bold": True, "border": 1, 'bg_color': '#DDDDDD'})
        fmtcpy = book.add_format({'align': 'left', "bold": True, "border": 1, 'bg_color': '#D98AF8'})
        fmtsec = book.add_format({'align': 'left', "bold": False, "border": 1, 'bg_color': '#66D248', 'font_size': 14})
        sheet = book.add_worksheet("Scrap Results")
        
        sheet.write_string("A1", "Overview", fmttbl)
        sheet.write_string("A2", "Name", fmttitl)
        sheet.write_string("B2", "Data", fmttitl)
        sheet.merge_range("A1:B1", "Overview", fmttbl)

        sheet.write_string("C1", "Officers", fmttbl)
        sheet.write_string("C2", "Number", fmttitl)
        sheet.write_string("D2", "Name", fmttitl)
        sheet.write_string("E2", "Status", fmttitl)
        sheet.write_string("F2", "Occupation", fmttitl)
        sheet.write_string("G2", "Role", fmttitl)
        sheet.write_string("H2", "Birth Date", fmttitl)
        sheet.write_string("I2", "Nationality", fmttitl)
        sheet.write_string("J2", "Address", fmttitl)
        sheet.write_string("K2", "Residence", fmttitl)
        sheet.write_string("L2", "Appointed/Resign", fmttitl)
        sheet.merge_range("C1:L1", "Officers", fmttbl)

        sheet.write_string("M1", "Histories", fmttbl)
        sheet.write_string("M2", "Number", fmttitl)
        sheet.write_string("N2", "Date", fmttitl)
        sheet.write_string("O2", "Info", fmttitl)
        sheet.write_string("P2", "URL", fmttitl)
        sheet.merge_range("N1:P1", "Histories", fmttbl)
        sheet.set_column(0, 20, 18)
        index = 3
        array = output['results']
        for node in array:
            pos = str(index)
            title = str(node['header']['companies']) + " companies found for \"" + node['header']['keywords'] + "\""
            sheet.write_string("A" + pos, title, fmtsec)
            sheet.merge_range("A" + pos + ":" + "P" + pos, title, fmtsec)
            index += 1
            data = node['matches']
            for record in data:
                pos = str(index)
                title = "[" + str(record['number']) + "] " + record['company'] + " (" + record['identity'] + ")"
                sheet.write_string("A" + pos, title, fmtcpy)
                sheet.merge_range("A" + pos + ":" + "P" + pos, title, fmtcpy)
                index += 1
                pos = str(index)
                over = record['overview']
                sheet.write_string("A" + pos, "Address", fmtinfo)
                sheet.write_string("B" + pos, over['address'], fmtinfo)
                pos = str(index + 1)
                sheet.write_string("A" + pos, "Status", fmtinfo)
                sheet.write_string("B" + pos, over['status'], fmtinfo)
                pos = str(index + 2)
                sheet.write_string("A" + pos, "Type", fmtinfo)
                sheet.write_string("B" + pos, over['type'], fmtinfo)
                pos = str(index + 3)
                sheet.write_string("A" + pos, "Incorporated", fmtinfo)
                sheet.write_string("B" + pos, over['incorporated'], fmtinfo)
                pos = str(index)
                hgover = 4
                hghist = 0
                hgoffc = 0
                users = record.get('officers', None)
                if not users is None:
                    for user in users:
                        sheet.write_string("C" + pos, str(hgoffc + 1), fmtinfo)
                        sheet.write_string("D" + pos, user.get('name', '-'), fmtinfo)
                        sheet.write_string("E" + pos, user.get('status', '-'), fmtinfo)
                        sheet.write_string("F" + pos, user.get('occupation', '-'), fmtinfo)
                        sheet.write_string("G" + pos, user.get('role', '-'), fmtinfo)
                        sheet.write_string("H" + pos, user.get('birth', '-'), fmtinfo)
                        sheet.write_string("I" + pos, user.get('nationality', '-'), fmtinfo)
                        sheet.write_string("J" + pos, user.get('address', '-'), fmtinfo)
                        sheet.write_string("K" + pos, user.get('residence', '-'), fmtinfo)
                        sheet.write_string("L" + pos, user.get('appointed', user.get('resigned', '-')), fmtinfo)
                        hgoffc += 1
                        pos = str(index + hgoffc)
                hist = record.get('histories', None)
                if not hist is None:
                    for info in hist:
                        sheet.write_string("M" + pos, str(info.get('no', '-')), fmtinfo)
                        sheet.write_string("N" + pos, info.get('date', '-'), fmtinfo)
                        sheet.write_string("O" + pos, info.get('desc', '-'), fmtinfo)
                        sheet.write_string("P" + pos, info.get('docs', '-'), fmtinfo)
                        hghist += 1
                        pos = str(index + hghist)
                index += max([hghist, hgoffc, hghist])        
        book.close()
        return path


print("> Initializing web scraper, please wait..")
settings = ScrapSettings()
print("> Loading web scraper settings..")
settings.reload()
print("> Initializing web scraper engine..")
provider = ScrapProvider(settings)
print("> Engine ready, starting scrap..")
try:
    output = provider.dispatch()
    print("> Operation success..")
    print("> Result with format JSON has saved: " + output[0])
    print("> Result with format XLSX has saved: " + output[1])
except Exception as e:
    print("> Scraping Error: " + str(e))