import requests
import time
import getpass
import os
import json
import html2text
import re
import shelve
import datetime
import xlsxwriter
import random
import string
from requests_toolbelt import MultipartEncoder

IHS_pages = {'Afghanistan': '1599312', 'Albania': '1600303',
             'Algeria': '1598638', 'American Samoa': '1598228',
             'Andorra': '1598208', 'Angola': '1598943', 'Anguilla': '1598181',
             'Antigua and Barbuda': '1600866', 'Argentina': '1598512',
             'Armenia': '1598200', 'Aruba': '1598220', 'Australia': '1598588',
             'Austria': '1600875', 'Azerbaijan': '1600269', 'Bahamas': '1598951',
             'Bahrain': '1599865', 'Bangladesh': '1599535', 'Barbados': '1600809',
             'Belarus': '1600720', 'Belgium': '1600245', 'Belize': '1600731',
             'Benin': '1599845', 'Bermuda': '1598235', 'Bhutan': '1600568',
             'Bolivia': '1599588', 'Bosnia and Herzegovina': '1600997',
             'Botswana': '1598203', 'Brazil': '1600094', 'Brunei': '1598225',
             'Bulgaria': '1598609', 'Burkina Faso': '1600949', 'Burundi': '1599991',
             'Cambodia': '1600937', 'Cameroon': '1599172', 'Canada': '1599023',
             'Cape Verde': '1598241', 'CAR': '1600062', 'Cayman Islands': '1598244',
             'Chad': '1599488', 'Chile': '1598683', 'China': '1599040',
             'Colombia': '1598617', 'Comorros': '1598270', 'Congo': '1598180',
             'Costa Rica':'1599362','Cote d\'Ivoire':'1598912','Croatia':'1599444',
             'Cuba':'1599679','Curacao':'1601147','Cyprus':'1599244',
             'Czech Republic':'1600932','Denmark':'1599251','Djibouti':'1600438',
             'Dominica':'1598162','Dominican Republic':'1600055','DR Congo':'1600389',
             'East Timor':'1600684','Ecuador':'1598941','Egypt':'1598767',
             'El Salvador':'1599848','Equatorial Guinea':'1598273','Eritrea':'1600642',
             'Estonia':'1600796','Ethiopia':'1598875','Fiji':'1599020','Finland':'1599963',
             'France':'1599140','French Guiana':'1600518','Gabon':'1599178',
             'Gambia':'1598276','Georgia':'1600273','Germany':'1599560','Ghana':'1598776','Greece':'1598657','Grenada':'1600793','Guam':'1598222','Guatemala':'1599711','Guinea':'1600023','Guinea-Bissau':'1598296','Guyana':'1600672','Haiti':'1599315','Honduras':'1598190','Hong Kong':'1598193','Hungary':'1599898','Iceland':'1599801','India':'1598878','Indonesia':'1598592','Iran':'1598797','Iraq':'1599468','Ireland':'1598196','Israel':'1599031','Italy':'1598663','Jamaica':'1600097','Japan':'1598470','Jordan':'1600178','Kazakhstan':'1599384','Kenya':'1598477','Kiribati':'1598299','Kosovo':'1600606','Kuwait':'1600402','Kyrgyzstan':'1601127','Laos':'1598215','Latvia':'1600221','Lebanon':'1599395','Lesotho':'1598292','Liberia':'1600621','Libya':'1599697','Liechtenstein':'1598187','Lithuania':'1600903','Luxembourg':'1600894','Macao':'1598277','Macedonia':'1601064','Madagascar':'1598289','Malawi':'1598286','Malaysia':'1600540','Maldives':'1599590','Mali':'1600513','Malta':'1600858','Marshall Islands':'1601418','Martinique':'1600644','Mauritania':'1601090','Mauritius':'1598268','Mexico':'1598845','Micronesia':'1598283','Moldova':'1600588','Mongolia':'1598764','Montenegro':'1600708','Morocco':'1599084','Mozambique':'1599371','Myanmar':'1599581','Namibia':'1598174','Nauru':'1601404','Nepal':'1600404','Netherlands':'1599894','New Zealand':'1598905','Nicaragua':'1600601','Niger':'1598280','Nigeria':'1598565','North Korea':'1599452','Norway':'1599547','Oman':'1600359','Pakistan':'1598632','Palau':'1601408','Palestine':'1600706','Panama':'1598673','Papua New Guinea':'1598980','Paraguay':'1599027','Peru':'1599368','Philippines':'1599957','Poland':'1599754','Portugal':'1600075','Puerto Rico':'1598218','Qatar':'1600279','Reunion':'1598259','Romania':'1600985','Russia':'1598339','Rwanda':'1599694','Saint Kitts and Nevis':'1598238','Saint Lucia':'1598251','Saint Vincent and the Grenadines':'1598248','Samoa':'1598256','Sao Tome and Principe':'1598266','Saudi Arabia':'1598602','Senegal':'1599714','Serbia':'1598361','Seychelles':'1598263','Sierra Leone':'1600213','Singapore':'1599430','Sint Maarten':'1601368','Slovakia':'1600521','Slovenia':'1598920','Solomon Islands':'1599874','Somalia':'1600958','South Africa':'1598954','South Korea':'1599366','South Sudan':'1600088','Spain':'1599175','Sri Lanka':'1598992','Sudan':'1600317','Suriname':'1600285','Swaziland':'1598184','Sweden':'1600855','Switzerland':'1601021','Syrian Arab Republic':'1599745','Taiwan':'1600216','Tajikistan':'1600969','Tanzania':'1599118','Thailand':'1599344','Togo':'1598177','Tonga':'1598232','Trinidad and Tobago':'1600071','Tunisia':'1599354','Turkey':'1598686','Turkmenistan':'1601018','Tuvalu':'1598229','UAE':'1599238','Uganda':'1599462','UK':'1598983','Ukraine':'1598680','Uruguay':'1600180','US Virgin Islands':'1598211','USA':'1598635','Uzbekistan':'1600487','Vanuatu':'1600435','Venezuela':'1599081','Vietnam':'1600701','Yemen':'1600964','Zambia':'1600065','Zimbabwe':'1600372'}


class LoadError(Exception):
    pass


def IHS_information():
    """
    Function is intended to gather information on country ratings (numbers) and
    country ratings information (words).  Returns 2 dictionaries with ratings
    and table information.
    """
    IHS_SSO_SESS, REMEMBER_ME = IHS_login()
    status, data = get_risk_ratings(IHS_SSO_SESS, REMEMBER_ME)
    if status == 200:
        rating_scores = ratings_parser(data)
    else:
        raise LoadError("Ratings failed to load")
    table_info = get_table_info()
    return rating_scores, table_info


def ratings_parser(data):
    """
    Function takes the received JSON and converts in dictionary with ratings information.
    """
    category_titles = []
    for line in data['Metadata']['Categories']:
        category_titles.append(line['LongHeader'])
    rating_scores = {}
    for line in data['Rows']:
        order = 0
        name = country_converter(line['Descriptor']['Country']['Name'])
        rating_scores[name] = {}
        for category in category_titles:
            if len(line['Cells'][order]) != 0:
                rating_scores[name][category] = line['Cells'][order]['Value']
            else:
                rating_scores[name][category] = None
            order += 1
    return rating_scores


def country_converter(name):
    """
    Function converts IHS country names into country names used by JTI.
    """
    converter = {'Curaçao': 'Curacao', 'Macao SAR': 'Macao', 'Czechia': 'Czech Republic',
                 'St Maarten': 'Sint Maarten', 'St Vincent and the Grenadines': 'Saint Vincent and the Grenadines',
                 'St Kitts and Nevis': 'Saint Kitts and Nevis', 'St Lucia': 'Saint Lucia',
                 'Virgin Islands, U.S.': 'US Virgin Islands', 'Hong Kong SAR': 'Hong Kong',
                 'Korea, South': 'South Korea', 'São Tomé and Príncipe': 'Sao Tome and Principe',
                 'Timor-Leste': 'East Timor', 'China (mainland)': 'China',
                 'Macedonia, North': 'Macedonia', 'United States': 'USA',
                 'United Kingdom': 'UK', 'Eswatini': 'Swaziland', 'Gambia, The': 'Gambia',
                 'United Arab Emirates': 'UAE', 'Congo, Republic of the': 'Congo',
                 'Korea, North': 'North Korea', 'Comoros': 'Comorros', 'Côte d\'Ivoire': 'Cote d\'Ivoire',
                 'Congo, the Democratic Republic of the': 'DR Congo', 'Palestinian Territory': 'Palestine',
                 'Central African Republic': 'CAR', 'Syria': 'Syrian Arab Republic'}
    if name in converter:
        return converter[name]
    else:
        return name


def IHS_login():
    """
    Function asks for user name and password and returns login credentials.
    """
    if os.path.isfile('IHS_credentials'):
        with open("IHS_credentials", 'r') as file:
            credentials = json.loads(file.read())
        time_dif = time.time() - credentials['Timestamp']
    else:
        time_dif = 60*60*24
    if time_dif < 60*60*24:
        return credentials['IHS_SSO_SESS'], credentials['REMEMBER_ME']
    else:
        url = "https://my.ihs.com/Connect/Login?ForceLogin=True"
        true = True
        username = input("Enter your IHS username: ")
        if username.endswith("@jti.com") is False:
            username += "@jti.com"
        password = getpass.getpass("Enter your password: ")
        payload = {"UserName": f"{username}", "Password": f"{password}", "RememberMe": true}
        webpage = requests.post(url, data=payload)
        if webpage.status_code == 200:
            IHS_SSO_SESS = webpage.json()["Session"]
            REMEMBER_ME = webpage.json()["Credential"]
            with open("IHS_credentials", 'w') as file:
                json.dump({'IHS_SSO_SESS': IHS_SSO_SESS, 'REMEMBER_ME': REMEMBER_ME,
                           'Timestamp': time.time()}, file)
            return IHS_SSO_SESS, REMEMBER_ME
        else:
            print("\nLogin credentials are incorrect")
            return None


def get_risk_ratings(IHS_SSO_SESS, REMEMBER_ME):
    """
    Function that gathers a status code and json file from IHS website.
    """
    params = {"IHS_SSO_SESS": f"{IHS_SSO_SESS}", "IHS_SSO_UI": f"{REMEMBER_ME}", "redirectUrl": "https://connect.ihs.com", "callingUrl": "https://my.ihs.com/Connect?ForceLogin=True", "theme": "Connect"}
    header = {"Cookie": "AMCV_92221CFE533057500A490D45%40AdobeOrg=359503849%7CMCIDTS%7C18516%7CMCMID%7C54559309136426318542700769955019031084%7CMCAAMLH-1600338999%7C6%7CMCAAMB-1600338999%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1599741400s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C5.0.1; _ga=GA1.2.1638942917.1599734202; mbox=session#08da7d772dee4731b7e20c60d6ac3f0f#1599736096|PC#08da7d772dee4731b7e20c60d6ac3f0f.37_0#1662979001; OptanonConsent=isIABGlobal=false&datestamp=Thu+Sep+10+2020+13%3A37%3A16+GMT%2B0300+(Moscow+Standard+Time)&version=5.11.0&landingPath=NotLandingPage&groups=C0001%3A1%2CC0002%3A1%2CC0003%3A1%2CC0004%3A1&hosts=bqo%3A1%2Cjwc%3A1%2Cneu%3A1%2Cqmo%3A1%2Clhz%3A1%2Cqwl%3A1%2Cjnl%3A1%2Cqag%3A1%2Ceda%3A1%2Cuar%3A1%2Chpo%3A1%2Ciad%3A1%2Civw%3A1%2Cfyg%3A1%2Cwit%3A1%2Cjjk%3A1%2Ckus%3A1%2Cqec%3A1%2Cyon%3A1%2Cnnh%3A1%2Cbmm%3A1%2Ckbb%3A1%2Clqg%3A1%2Ctgi%3A1%2Ccqb%3A1%2Cqag%3A1%2Czom%3A1%2Ceda%3A1%2Cqrr%3A1%2Chpo%3A1%2Cudv%3A1%2Czdk%3A1%2Csln%3A1%2Cpgx%3A1%2Cwhx%3A1%2Ceff%3A1%2Cneu%3A1%2Ctxw%3A1%2Ckeu%3A1%2Cbys%3A1%2Cvvv%3A1%2Cqqh%3A1%2Clhz%3A1%2Cnon%3A1%2Cyio%3A1%2Cikx%3A1&AwaitingReconsent=false; ihsmarkit=a07bf3ca.5b186b6e6c1d9; IHS_SSO_UI={}; _gid=GA1.2.1819444186.1602567041; IHS_SSO_SESS={}".format(REMEMBER_ME, IHS_SSO_SESS)}
    ratings_url = f"https://connect.ihsmarkit.com/Risks/GetValues?Service=CountryRiskBeta%23SecurityPlanner&Date=&SortColumn=Rank&SortOrder=asc&_={int(time.time())}"
    ratings = requests.get(ratings_url, data=params, headers=header)
    return ratings.status_code, ratings.json()


def get_table_info():
    """
    Function checks pages of every country in the list on IHS website.
    Returns a dictionary with table information.
    """
    table_dictionary = {}
    for page in IHS_pages:
        print("Loading... {}".format(page))
        searchable = get_html(IHS_pages[page])
        try:
            search_result = replace_new_lines(re.search("Security[\n]([\d\D]*)", searchable).group(1))
            search_result = search_result.split('\n')
            table_dictionary[page] = gather_risks(search_result)
        except AttributeError:
            table_dictionary[page] = {}
    return table_dictionary


def replace_new_lines(result):
    line = "QWERTYUIOPASDFGHJKLZXCVBNMqweértyuiopasdfghjklzxcvbnm1234567890,:;–"
    for letter in line:
        for letter2 in line:
            new_result = re.sub('{}\n{}'.format(letter, letter2), '{} {}'.format(letter, letter2), result)
            result = new_result
        for letter in line:
            new_result = re.sub('\)\n{}'.format(letter), '\) {}'.format(letter), result)
            result = new_result
        for letter in line:
            new_result = re.sub('{}\n\('.format(letter), '{} \('.format(letter), result)
            result = new_result
        for letter in line:
            new_result = re.sub('\]\n{}'.format(letter), '\] {}'.format(letter), result)
            result = new_result
        for letter in line:
            new_result = re.sub('{}\n\['.format(letter), '{} \['.format(letter), result)
            result = new_result
        for letter in line:
            new_result = re.sub('-\n{}'.format(letter), '- {}'.format(letter), result)
            result = new_result
    return result


def gather_risks(search_result):
    risks_dict = {}
    risks = ["War risks", "Terrorism risks", "Social stability and unrest risks",
             "Risks to individuals", "Risks to cargo/transport", "Risks to property"]
    for line in search_result:
        if line in risks:
            risks_dict[line] = []
            position = search_result.index(line)
            for num in range(position+2, len(search_result)):
                if search_result[num] == "" or search_result[num] in risks:
                    break
                else:
                    risks_dict[line].append(search_result[num])
    return risks_dict


def get_html(page_number):
    IHS_SSO_SESS, REMEMBER_ME = IHS_login()
    params = {"IHS_SSO_SESS": f"{IHS_SSO_SESS}", "IHS_SSO_UI": f"{REMEMBER_ME}", "redirectUrl": "https://connect.ihs.com", "callingUrl": "https://my.ihs.com/Connect?ForceLogin=True", "theme": "Connect"}
    header = {"Cookie": "AMCV_92221CFE533057500A490D45%40AdobeOrg=359503849%7CMCIDTS%7C18516%7CMCMID%7C54559309136426318542700769955019031084%7CMCAAMLH-1600338999%7C6%7CMCAAMB-1600338999%7CRKhpRz8krg2tLO6pguXWp5olkAcUniQYPHaMWWgdJ3xzPWQmdj0y%7CMCOPTOUT-1599741400s%7CNONE%7CMCAID%7CNONE%7CvVersion%7C5.0.1; _ga=GA1.2.1638942917.1599734202; mbox=session#08da7d772dee4731b7e20c60d6ac3f0f#1599736096|PC#08da7d772dee4731b7e20c60d6ac3f0f.37_0#1662979001; OptanonConsent=isIABGlobal=false&datestamp=Thu+Sep+10+2020+13%3A37%3A16+GMT%2B0300+(Moscow+Standard+Time)&version=5.11.0&landingPath=NotLandingPage&groups=C0001%3A1%2CC0002%3A1%2CC0003%3A1%2CC0004%3A1&hosts=bqo%3A1%2Cjwc%3A1%2Cneu%3A1%2Cqmo%3A1%2Clhz%3A1%2Cqwl%3A1%2Cjnl%3A1%2Cqag%3A1%2Ceda%3A1%2Cuar%3A1%2Chpo%3A1%2Ciad%3A1%2Civw%3A1%2Cfyg%3A1%2Cwit%3A1%2Cjjk%3A1%2Ckus%3A1%2Cqec%3A1%2Cyon%3A1%2Cnnh%3A1%2Cbmm%3A1%2Ckbb%3A1%2Clqg%3A1%2Ctgi%3A1%2Ccqb%3A1%2Cqag%3A1%2Czom%3A1%2Ceda%3A1%2Cqrr%3A1%2Chpo%3A1%2Cudv%3A1%2Czdk%3A1%2Csln%3A1%2Cpgx%3A1%2Cwhx%3A1%2Ceff%3A1%2Cneu%3A1%2Ctxw%3A1%2Ckeu%3A1%2Cbys%3A1%2Cvvv%3A1%2Cqqh%3A1%2Clhz%3A1%2Cnon%3A1%2Cyio%3A1%2Cikx%3A1&AwaitingReconsent=false; ihsmarkit=a07bf3ca.5b186b6e6c1d9; IHS_SSO_UI={}; _gid=GA1.2.1819444186.1602567041; IHS_SSO_SESS={}".format(REMEMBER_ME, IHS_SSO_SESS)}
    response_url = f"https://connect.ihsmarkit.com/Document/Show/phoenix/{page_number}?connectPath=CountryRisk.CRptCountryReportWidget&_=1602581825861"
    response_page = requests.get(response_url, data=params, headers=header)
    if response_page.status_code != 200:
        raise LoadError("Table information failed to load")
    return html2text.html2text(response_page.text)


def gather_current_ratings(object_type_id):
    """
    Function logs into SIMP and gathers all IHS ratings that are active.
    Returns a list of tuples with names and external ref ids.
    """
    headers = {"Authorization": f"bearer {authentication()}"}
    pages = count_objects(object_type_id) + 1
    final_list = []
    for page in range(1, pages):
        retrieve_all_objects_url = f"https://eu.core.resolver.com/data/object?objectTypeId={object_type_id}&pageNumber={page}"
        all_objects = requests.get(retrieve_all_objects_url, headers=headers).json()
        final_list.extend(parse_objects(all_objects['data']))
    return final_list


def count_objects(object_type_id):
    """
    Functions logs into SIMP and returns the count of pages for objects in SIMP.
    """
    headers = {"Authorization": f"bearer {authentication()}"}
    count_url = f"https://eu.core.resolver.com/data/object/count?objectTypeId={object_type_id}"
    counted_objects = requests.get(count_url, headers=headers).text
    pages = int(counted_objects) // 100 + 1
    return pages


def parse_objects(all_objects):
    """
    Converts the dictionary from SIMP into tuples with name and externalRefId.
    """
    resulting_list = []
    for object_ in all_objects:
        if object_['objectLifeCycleStateId'] == 34293:
            resulting_list.append((object_['name'], object_['externalRefId']))
    return resulting_list


def authentication():
    """
    Function used for SIMP authentication
    """
    try:
        shelf_file = shelve.open('production_token_info')
        shelf_file['expiration']
    except KeyError:
        shelf_file['expiration'] = 0
        shelf_file.close()
        pass
    shelf_file = shelve.open('production_token_info')
    if time.time() < shelf_file['expiration']:
        token = shelf_file['token']
        return token
    else:
        username = input("Enter your SIMP username: ")
        if username.endswith("@jti.com") is False:
            username += "@jti.com"
        password = getpass.getpass("Enter your password: ")
        authenticate_url = 'https://eu.core.resolver.com/user/authenticate'
        data = {"email": username, "password": password,
                "selectedOrg": 166, "client": "core-client"}
        login_to_sandbox = requests.post(authenticate_url, json=data)
        token_information = login_to_sandbox.json()
        try:
            token = shelf_file['token'] = token_information['token']
            shelf_file['expiration'] = token_information['expiresAt']
            shelf_file.close()
            return token
        except KeyError:
            print("\nLogin credentials are incorrect")
            return None


def create_excel_file(scores, table_info, active_ratings):
    """
    Function takes current IHS scores, information from tables and active ratings refIds.
    Saves a file to the same folder as script.  Returns filename.
    """
    # Creating Worksheet
    today = datetime.date.today().strftime("%Y-%m-%d")
    excel_date = datetime.datetime.strptime(today, "%Y-%m-%d").date()
    file_name = "IHS_upload_file_" + today + ".xlsx"
    new_file = xlsxwriter.Workbook(file_name)
    # First sheet in workbook - Setting relationships with new IHS ratings
    sheet1 = new_file.add_worksheet('CNT - IHS - Country Summary Re')
    titles = [("Relationship ID", 0, 0), ("7bcc5c5a-5045-4c69-8a7c-885d544782ef", 1, 0),
              ("Country",2,0), ("OB1 Ext Ref ID",3,0), ("Object Type ID",0,1),
              ("Country", 1, 1), ("(optional)", 2, 1), ("Object Name", 3, 1),
              ("Object Type ID",0,2), ("3c9e1b5b-99a8-406a-ab0b-7ba48b9a9ff6",1,2),
              ("IHS - Country Summary Report",2,2), ("OB2 Ext Ref ID",3,2),
              ("(optional)",2,3), ("Object Name",3,3)]
    for item in titles:
        name, line, column = item
        sheet1.write(line, column, name)
    line = 4
    for country in scores:
        if country in IHS_pages:
            sheet1.write(line, 0, country)
            sheet1.write(line, 1, country)
            report_name = "IHS: " + country + " - " + today
            sheet1.write(line, 2, report_name)
            sheet1.write(line, 3, report_name)
            line += 1

    # Second sheet in workbook - Main sheet with all the information
    sheet2 = new_file.add_worksheet('IHS - Country Summary Report')
    sheet2.write(0, 0, "Object Type ID")
    sheet2.write(1, 0, "3c9e1b5b-99a8-406a-ab0b-7ba48b9a9ff6")
    sheet2_titles = ['Risk to Cargo/Transport ', 'Crime', 'Protests',
                     'Terrorism  Risks ', 'Interstate War ', 'Cargo: Aviation Risks',
                     'Business Risk rating:', 'Overall Cargo', 'Civil War - Rating',
                     'Crime - Rating', 'Cargo: Ground Risks', 'Interstate War - Rating',
                     'Kidnap - Rating', 'Cargo: Maritime Risks', 'Protests - Rating',
                     'Terrorism - Rating', 'Date Logged', 'Latest Date', 'Library Workflow']
    core_titles = ['External Ref ID', 'Name', 'Description', 'RISKTOCARG', 'CRIME',
                   'PROTESTS', 'TERRORISMR', 'INTERSTATE', 'CARGO:AVIA~1',
                   'IHS-OPERAT', 'IHS-CARGOR', 'IHS-CIVILW', 'CRIME-RATI',
                   'CARGO:GROU~1', 'IHS-INTERS', 'IHS-KIDNAP', 'CARGO:MARI~1',
                   'IHS-PROTES', 'IHS-TERROR', 'DATELOGGED', 'LATESTDATE',
                   'IHS - Country Summary Report']
    column = 3
    for title in sheet2_titles:
        sheet2.write(2, column, title)
        column += 1
    column = 0
    for title in core_titles:
        sheet2.write(3, column, title)
        column += 1
    # FILL OUT THE SHEET
    line = 4
    for country in table_info:
        sheet2.write(line, 0, f"IHS: {country} - {today}")
        sheet2.write(line, 1, f"IHS: {country} - {today}")
        sheet2.write(line, 3, exract_info(table_info[country], "Risks to cargo/transport"))
        sheet2.write(line, 4, exract_info(table_info[country], "Risks to individuals"))
        sheet2.write(line, 5, exract_info(table_info[country], "Social stability and unrest risks"))
        sheet2.write(line, 6, exract_info(table_info[country], "Terrorism risks"))
        sheet2.write(line, 7, exract_info(table_info[country], "War risks"))
        sheet2.write(line, 8, extract_scores(scores[country], "Aviation"))
        sheet2.write(line, 9, extract_scores(scores[country], "Business risk"))
        sheet2.write(line, 10, extract_scores(scores[country], "Cargo and transport"))
        sheet2.write(line, 11, extract_scores(scores[country], "Civil war"))
        sheet2.write(line, 12, extract_scores(scores[country], "Criminal violence"))
        sheet2.write(line, 13, extract_scores(scores[country], "Ground"))
        sheet2.write(line, 14, extract_scores(scores[country], "Interstate war"))
        sheet2.write(line, 15, extract_scores(scores[country], "Kidnap and ransom"))
        sheet2.write(line, 16, extract_scores(scores[country], "Marine"))
        sheet2.write(line, 17, extract_scores(scores[country], "Protests and riots"))
        sheet2.write(line, 18, extract_scores(scores[country], "Terrorism"))
        sheet2.write_datetime(line, 19, excel_date)
        sheet2.write_datetime(line, 20, excel_date)
        sheet2.write(line, 21, "Active")
        line += 1
    for object_ in active_ratings:
        sheet2.write(line, 0, object_[0])
        sheet2.write(line, 1, object_[1])
        sheet2.write(line, 21, "Archived")
        line += 1
    new_file.close()
    return file_name


def exract_info(dictionary, risk_name):
    if risk_name in dictionary:
        return " ".join(dictionary[risk_name])
    else:
        return ""


def extract_scores(dictionary, score_name):
    if dictionary[score_name] is None:
        return ""
    score = float(dictionary[score_name])
    print(score - int(score))
    if score - int(score) == 0:
        score = int(score)
    if score <= 0.7:
        return str(score) + " LOW"
    elif score > 0.7 and score <= 1.5:
        return str(score) + " MODERATE"
    elif score > 1.5 and score <= 2.3:
        return str(score) + " ELEVATED"
    elif score > 2.3 and score <= 3.1:
        return str(score) + " HIGH"
    elif score > 3.1 and score <= 4.3:
        return str(score) + " VERY HIGH"
    elif score > 4.3 and score <= 6.4:
        return str(score) + " SEVERE"
    elif score > 6.4 and score <= 10:
        return str(score) + " EXTREME"


def send_file_to_SIMP(file_name):
    """
    Takes the file, uplods it and pases to SIMP, return a job ID
    """
    false, true = False, True
    upload_url = "https://eu.core.resolver.com/creation/import"
    data = open(file_name, 'rb')
    fields = {'file': (file_name, data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
    m = MultipartEncoder(fields=fields, boundary='----WebKitFormBoundary'+''.join(random.sample(string.ascii_letters+string.digits, 16)))
    token = authentication()
    headers = {"Connection": "keep-alive", "Content-Type": m.content_type,
               "Authorization": f"bearer {token}"}
    payload = {"dryRun": false, "usingExternalRefIds": true,
               "deferPostProcessing": false}
    uploading = requests.post(upload_url, headers=headers, params=payload, data=m)
    return uploading.json()['jobId']


def checking_job(job_id):
    """
    Logs into SIMP and checks the status of upload.
    """
    status = 1
    while status == 1:
        token = authentication()
        header = {'Authorization': 'bearer {}'.format(token)}
        poll_status_url = f"https://eu.core.resolver.com/object/job/{job_id}"
        poll_status = requests.get(poll_status_url, headers=header)
        status = poll_status.json()['status']
        print("... Checking job status")
        time.sleep(15)
    ts = int(poll_status.json()['finished'])
    end_time = datetime.datetime.utcfromtimestamp(ts).strftime('%Y-%m-%d %H:%M:%S')
    return end_time


def main():
    """
    Main function.
    """
    checking_IHS_login = None
    while checking_IHS_login is None:
        checking_IHS_login = IHS_login()
    checking_SIMP_login = None
    while checking_SIMP_login is None:
        checking_SIMP_login = authentication()
    raise Exception
    scores, table_info = IHS_information()
    active_ratings = gather_current_ratings(8891)
    file_name = create_excel_file(scores, table_info, active_ratings)
    job_id = send_file_to_SIMP(file_name)
    upload_time = checking_job(job_id)
    print(f"File uploaded at {upload_time} UTC")


if __name__ == "__main__":
    main()
