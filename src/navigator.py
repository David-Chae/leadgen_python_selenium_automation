from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from nameparser import HumanName
import xlsxwriter
import time
import re, string
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException


titles = ('Chief Marketing Officer')

#companies = ('Allotrac', 'Fulton Market Group', 'Board International', 'Servigistics, a PTC Technology', 'ToolsGroup', 'Bodd','Department of Transport', 'Colliers', 'Caltex Australia', 'Findex', 'BHD Storage Solutions', 'Retailquip Pty Ltd', 'Scanreco Group', 'DriveRisk (Pty) Ltd', 'TransVirtual', 'Teletrac Navman New Zealand', 'Merit Manufacturing', 'Kynection','Toowoomba Regional Council', 'Sadleirs', 'Multi-Mover Europe BV', 'Flinders Ports Pty Limited', 'DEAN World Cargo', 'Pathtech Pty Ltd', 'Diverseco', 'MTData', 'PTV Group IMEA', 'Dematic Mobile Automation', 'Dematic', 'Headland Machinery', 'DriveRisk Australasia', 'Transurban', 'Asafe Greece - Inoxtec', 'LOSCAM Group', 'Simoco Wireless Solutions', 'Combilift', 'Australian Border Force', 'PTV Group Asia Pacific', 'Sweepers Australia', 'Gilbarco Veeder-Root', 'PTV Group', 'SSAB','McNaughtans Pty. Ltd.' , 'ToooAir Pty. Ltd.', 'Hard Surface Cleaners - Pressure Cleaning & Soft Washing', 'Gaprie Ltd', 'Damon Technology', 'VISA GLOBAL LOGISTICS NL', 'Cascade Australia', 'eCycle Solutions', 'GoLive Monitoring', 'Precision Automotive Equipment', 'Promata Automotive', 'Direct Mail Corporaton', 'Multi-Mover UK', 'Gilbarco Veeder-Root MEA', 'Blue Jay Solutions', 'EForklift Pty Ltd', 'vWork', 'Protection Experts Australia', 'Tele Radio Group', 'MTDATA LIMITED', 'Easy Wash Australia', 'Supply Chain Logistics Association of Australia (SCLAA)', 'Teletrac Navman', 'MTData LLC', 'MTData NZ', 'Automation Systems & Controls Pty Ltd', 'Lionel Samson Sadleirs Group', 'Peacock Bros.', 'McLardy McShane Insurance Brokers', 'R&R Corporate Health', 'Combilift Depot', 'Conquest Equipment', 'Fischer Plastic Products Pty Ltd', 'Microlise', 'Envirofluid', 'Allotrac', 'ASCI - Australasian Supply Chain Institute','Truckworld.com.au', 'Australian Sweeper Company', 'Teletrac Navman Australia', 'KAB Seating Limited', 'Transficient - representing TK Blue Agency & Global Climate Initiatives in Australia', 'ANL', 'Leopard Systems', 'ULTIMATE LED', 'Goodyear Dunlop Tyres (Australia) Pty Ltd', 'Australia LOAD SMART TRUCKING INC', 'Transport Certification Australia', 'Transurban Group', 'Dematic Retrotech', 'Scaco Pty Ltd', 'Gilbarco Australia Holding Pty Ltd', 'Cincom Systems', 'Sadleirs Transport', 'Cotewell', 'Port of Newcastle', 'National Heavy Vehicle Regulator', 'ASAFE AUSTRALASIA Pty Ltd', 'X-Pak Global', 'Datanet Asia Pacific Pty Ltd', 'Ultimate LED Lights', 'WISTA International', 'WIM Technologies', 'ANC Distribution', 'Tele Radio Group Australia', 'Colliers Engineering Design', 'AI Drive By GoLive Monitoring', 'KBS, a Premier Tech Company', 'Netstar Australia', 'TriTech Lubricants Australasia', 'ANC Distribution Australia Pty Ltd', 'Eonmetall Official', 'Muscat Trailers', 'Combilift Australia', 'Bustle Tech', 'M2MONEY ONLINE SERVICE', 'Yarno', 'Gilbarco Veeder Root SEA', 'Infocomm Pty Ltd', 'AUSTRALIAN SWEEPER CORPORATION PTY LTD', 'Nilfisk', 'Bailey Ladders', 'Global GPS Tracking', 'GILBARCO')
company_list = ['Servigistics, a PTC Technology', 'VNet', 'Thinxtra, The IoT Telco', 'enVista', 'Board International', 'Dematic', 'Dematic Mobile Automation', 'Dematic Retrotech', 'Aera Technology', 'Anaplan', 'Basware', 'Basware Benelux', 'BCR Australia Pty Ltd', 'Blue Yonder', 'Bluejay Solutions', 'BlueJay Web Solutions', 'Coupa Software', 'Coupa Japan', 'Coupa Supply Chain, powered by LLamasoft', 'Coupa Suppliers', 'Descartes Systems Group', 'Descartes Systems Group France', 'Descartes Systems Group Benelux', 'Disprax ERP Software Australia', 'E2open', 'EPG Ehrhardt Partner Group', 'topsystem GmbH (Member of EPG)', 'Epicor Software', 'Garvis', 'GEP Worldwide', 'GEP APAC', 'Körber', 'Körber Supply Chain', 'Körber Supply Chain APAC', 'IBM', 'Icron | An Analog Devices Brand', 'IFS', 'Infocomm Pty Ltd', 'Infor', 'infor global solutions', 'Insight', 'Ivalua', 'JAGGAER', 'JDA Software', 'JLL', 'JLL Asia Pacific', 'Kinaxis', 'Lexian Solutions', 'Logility', 'Manhattan Associates', 'Optimity', 'Optimity Software', 'NetSuite', 'project44', 'SAP', 'SAP Customer Experience', 'SoftwareONE', 'SoftwareONE Australia', 'Sonatype', 'TMC, a division of C.H. Robinson', '| uTenant. | The Warehousing Matchmaker', 'V Net Solutions', 'WiseTech Global']



leads = []

page_urls = []

saved_search = "https://www.linkedin.com/sales/search/people?query=(recentSearchParam%3A(id%3A1346013636%2CdoLogHistory%3Atrue)%2Cfilters%3AList((type%3AREGION%2Cvalues%3AList((id%3A103313686%2Ctext%3ANew%2520South%2520Wales%252C%2520Australia%2CselectionType%3AINCLUDED)))%2C(type%3ACOMPANY_HEADCOUNT%2Cvalues%3AList((id%3AC%2Ctext%3A11-50%2CselectionType%3AINCLUDED)))%2C(type%3ASENIORITY_LEVEL%2Cvalues%3AList((id%3A3%2Ctext%3AEntry%2CselectionType%3AINCLUDED)))%2C(type%3ACOMPANY_HEADQUARTERS%2Cvalues%3AList((id%3A100506852%2Ctext%3AAustralia%2CselectionType%3AINCLUDED)))%2C(type%3AFUNCTION%2Cvalues%3AList((id%3A13%2Ctext%3AInformation%2520Technology%2CselectionType%3AINCLUDED)))))&sessionId=dgQfxmG7RT6qFb2Ku3ouGA%3D%3D"

#Priority group
saved_search1 = "https://www.linkedin.com/sales/search/people?companyIncluded=Allotrac%3A3575424%2CFulton%2520Market%2520Group%3A13461093%2CBoard%2520International%3A98246%2CServigistics%252C%2520a%2520PTC%2520Technology%3A13667%2CToolsGroup%3A32403%2CBodd%3A10701494&companyTimeScope=CURRENT&doFetchHeroCard=false&geoIncluded=101452733&logHistory=true&rsLogId=1322162068&searchSessionId=%2BmOR94sNQLCUtV5SZ8wFtw%3D%3D&seniorityIncluded=6%2C8%2C7"

#marketing top people
saved_search2 = "https://www.linkedin.com/sales/search/people#companyIncluded=Department%2520of%2520Transport%3A349778%2CColliers%3A5227%2CCaltex%2520Australia%3A9965%2CFindex%3A3967749%2CBHD%2520Storage%2520Solutions%3A927862%2CRetailquip%2520Pty%2520Ltd%3A10038173%2CScanreco%2520Group%3A10104011%2CDriveRisk%2520(Pty)%2520Ltd%3A10180765%2CTransVirtual%2520%3A10530800%2CTeletrac%2520Navman%2520%257C%2520New%2520Zealand%3A10691036%2CMerit%2520Manufacturing%3A10701511%2CKynection%3A1070283%2CToowoomba%2520Regional%2520Council%3A1075896%2CSadleirs%3A1086247%2CMulti-Mover%2520Europe%2520BV%3A11009737%2CFlinders%2520Ports%2520Pty%2520Limited%3A1111798%2CDEAN%2520World%2520Cargo%3A1166836%2CPathtech%2520Pty%2520Ltd%3A1171393%2CDiverseco%3A12906515%2CMTData%3A131705%2CPTV%2520Group%2520IMEA%2520%3A13348499%2CDematic%2520Mobile%2520Automation%3A136109%2CDematic%3A14074%2CHeadland%2520Machinery%3A1431001%2CDriveRisk%2520Australasia%3A14397521%2CTransurban%3A14488%2CAsafe%2520Greece%2520-%2520Inoxtec%2520%3A14807739%2CLOSCAM%2520Group%3A1483532%2CSimoco%2520Wireless%2520Solutions%3A148472%2CCombilift%3A1486399%2CAustralian%2520Border%2520Force%3A15309%2CPTV%2520Group%2520Asia%2520Pacific%3A16153489%2CSweepers%2520Australia%3A1629292%2CGilbarco%2520Veeder-Root%3A163991%2CPTV%2520Group%3A165593%2CSSAB%3A166109%2CMcNaughtans%2520Pty.%2520Ltd.%3A18098583%2CToooAir%2520Pty.%2520Ltd.%3A18099392%2CHard%2520Surface%2520Cleaners%2520-%2520Pressure%2520Cleaning%2520%2526%2520Soft%2520Washing%3A18137625%2CGaprie%2520Ltd%3A18313932%2CDamon%2520Technology%3A18349726%2CVISA%2520GLOBAL%2520LOGISTICS%2520NL%3A18459337%2CCascade%2520Australia%3A18534998%2CeCycle%2520Solutions%3A1860791%2CGoLive%2520Monitoring%3A18613723%2CPrecision%2520Automotive%2520Equipment%3A18711214%2CPromata%2520Automotive%3A18758236%2CDirect%2520Mail%2520Corporaton%3A18836595%2CMulti-Mover%2520UK%3A18893256%2CGilbarco%2520Veeder-Root%2520MEA%3A19087182%2CBlue%2520Jay%2520Solutions%3A19121476%2CEForklift%2520Pty%2520Ltd%3A19243240%2CvWork%3A1991796%2CProtection%2520Experts%2520Australia%3A20387540%2CTele%2520Radio%2520Group%3A2083055%2CMTDATA%2520LIMITED%3A21314192%2CEasy%2520Wash%2520Australia%3A22301968%2CSupply%2520Chain%2520%2526%2520Logistics%2520Association%2520of%2520Australia%2520(SCLAA)%3A2399221%2CTeletrac%2520Navman%3A24990%2CMTData%2520LLC%3A2629625%2CMTData%2520NZ%3A2632532%2CAutomation%2520Systems%2520%2526%2520Controls%2520Pty%2520Ltd%3A3066360%2CLionel%2520Samson%2520Sadleirs%2520Group%3A3114048%2CPeacock%2520Bros.%3A3219327%2CMcLardy%2520McShane%2520Insurance%2520Brokers%3A322227%2CR%2526R%2520Corporate%2520Health%3A3276421%2CCombilift%2520Depot%3A3314927%2CConquest%2520Equipment%3A3361929%2CFischer%2520Plastic%2520Products%2520Pty%2520Ltd%3A3364182%2CMicrolise%3A33839%2CEnvirofluid%3A3552049%2CAllotrac%3A3575424%2CASCI%2520-%2520Australasian%2520Supply%2520Chain%2520Institute%3A3682349%2CTruckworld.com.au%3A3692277%2CAustralian%2520Sweeper%2520Company%3A3738673%2CTeletrac%2520Navman%2520%257C%2520Australia%3A3742855%2CKAB%2520Seating%2520Limited%3A4027652%2CTransficient%2520-%2520representing%2520TK'Blue%2520Agency%2520%2526%2520Global%2520Climate%2520Initiatives%2520in%2520Australia%3A42697837%2CANL%3A43609%2CLeopard%2520Systems%3A486727%2CULTIMATE%2520LED%3A48964277%2CGoodyear%2520%2526%2520Dunlop%2520Tyres%2520(Australia)%2520Pty%2520Ltd%2520Australia%3A53770224%2CLOAD%2520SMART%2520TRUCKING%2520INC%3A53829130%2CTransport%2520Certification%2520Australia%3A545896%2CTransurban%2520Group%3A55416951%2CDematic%2520Retrotech%3A56675%2CScaco%2520Pty%2520Ltd%3A58555205%2CGilbarco%2520Australia%2520Holding%2520Pty%2520Ltd%3A58681899%2CCincom%2520Systems%3A6160%2CSadleirs%2520Transport%3A62647764%2CCotewell%3A6423068%2CPort%2520of%2520Newcastle%3A6432433%2CNational%2520Heavy%2520Vehicle%2520Regulator%3A6520823%2CASAFE%2520AUSTRALASIA%2520Pty%2520Ltd%3A6590925%2CX-Pak%2520Global%3A6626434%2CDatanet%2520Asia%2520Pacific%2520Pty%2520Ltd%3A662994%2CUltimate%2520LED%2520Lights%3A67776597%2CWISTA%2520International%2520%3A68796735%2CWIM%2520Technologies%3A69544956%2CANC%2520Distribution%2520%3A71013196%2CTele%2520Radio%2520Group%2520Australia%2520%3A71622572%2CColliers%2520Engineering%2520%2526%2520Design%3A71630305%2CAI%2520Drive%2520By%2520GoLive%2520Monitoring%3A71992449%2CKBS%252C%2520a%2520Premier%2520Tech%2520Company%3A749880%2CNetstar%2520Australia%3A757162%2CTriTech%2520Lubricants%2520Australasia%3A7616600%2CANC%2520Distribution%2520Australia%2520Pty%2520Ltd%3A7670405%2CEonmetall%2520Official%3A77077646%2CMuscat%2520Trailers%3A7750950%2CCombilift%2520Australia%3A7779401%2CBustle%2520Tech%3A7791457%2CM2MONEY%2520ONLINE%2520SERVICE%3A78910980%2CYarno%3A7965238%2CGilbarco%2520Veeder%2520Root%2520SEA%3A79842426%2CInfocomm%2520Pty%2520Ltd%3A8177974%2CAUSTRALIAN%2520SWEEPER%2520CORPORATION%2520PTY%2520LTD%3A8274156%2CNilfisk%3A9029%2CBailey%2520Ladders%3A9125623%2CGlobal%2520GPS%2520Tracking%3A9580664%2CGILBARCO%3A9739658&companyTimeScope=CURRENT&doFetchHeroCard=false&functionIncluded=15&geoIncluded=101452733&logHistory=true&rsLogId=1322161036&searchSessionId=tkTbQPkmTXiU807BqoHmDg%3D%3D&selectedFilter=GE&seniorityIncluded=6%2C7%2C8%2C5"

#MDs and GMs
saved_search3 = "https://www.linkedin.com/sales/search/people#companyIncluded=Colliers%3A5227%2CCaltex%2520Australia%3A9965%2CMcLardy%2520McShane%2520Insurance%2520Brokers%3A322227%2CNilfisk%3A9029%2CBHD%2520Storage%2520Solutions%3A927862%2CANC%2520Distribution%2520Australia%2520Pty%2520Ltd%3A7670405%2CYarno%3A7965238%2CInfocomm%2520Pty%2520Ltd%3A8177974%2CRetailquip%2520Pty%2520Ltd%3A10038173%2CScanreco%2520Group%3A10104011%2CDriveRisk%2520(Pty)%2520Ltd%3A10180765%2CTransVirtual%2520%3A10530800%2CTeletrac%2520Navman%2520%257C%2520New%2520Zealand%3A10691036%2CMerit%2520Manufacturing%3A10701511%2CKynection%3A1070283%2CToowoomba%2520Regional%2520Council%3A1075896%2CSadleirs%3A1086247%2CMulti-Mover%2520Europe%2520BV%3A11009737%2CFlinders%2520Ports%2520Pty%2520Limited%3A1111798%2CDEAN%2520World%2520Cargo%3A1166836%2CPathtech%2520Pty%2520Ltd%3A1171393%2CDiverseco%3A12906515%2CMTData%3A131705%2CPTV%2520Group%2520IMEA%2520%3A13348499%2CDematic%2520Mobile%2520Automation%3A136109%2CDematic%3A14074%2CHeadland%2520Machinery%3A1431001%2CDriveRisk%2520Australasia%3A14397521%2CTransurban%3A14488%2CAsafe%2520Greece%2520-%2520Inoxtec%2520%3A14807739%2CLOSCAM%2520Group%3A1483532%2CSimoco%2520Wireless%2520Solutions%3A148472%2CCombilift%3A1486399%2CAustralian%2520Border%2520Force%3A15309%2CPTV%2520Group%2520Asia%2520Pacific%3A16153489%2CSweepers%2520Australia%3A1629292%2CGilbarco%2520Veeder-Root%3A163991%2CPTV%2520Group%3A165593%2CSSAB%3A166109%2CMcNaughtans%2520Pty.%2520Ltd.%3A18098583%2CToooAir%2520Pty.%2520Ltd.%3A18099392%2CHard%2520Surface%2520Cleaners%2520-%2520Pressure%2520Cleaning%2520%2526%2520Soft%2520Washing%3A18137625%2CGaprie%2520Ltd%3A18313932%2CDamon%2520Technology%3A18349726%2CVISA%2520GLOBAL%2520LOGISTICS%2520NL%3A18459337%2CCascade%2520Australia%3A18534998%2CeCycle%2520Solutions%3A1860791%2CGoLive%2520Monitoring%3A18613723%2CPrecision%2520Automotive%2520Equipment%3A18711214%2CPromata%2520Automotive%3A18758236%2CDirect%2520Mail%2520Corporaton%3A18836595%2CMulti-Mover%2520UK%3A18893256%2CGilbarco%2520Veeder-Root%2520MEA%3A19087182%2CBlue%2520Jay%2520Solutions%3A19121476%2CEForklift%2520Pty%2520Ltd%3A19243240%2CvWork%3A1991796%2CProtection%2520Experts%2520Australia%3A20387540%2CTele%2520Radio%2520Group%3A2083055%2CMTDATA%2520LIMITED%3A21314192%2CEasy%2520Wash%2520Australia%3A22301968%2CSupply%2520Chain%2520%2526%2520Logistics%2520Association%2520of%2520Australia%2520(SCLAA)%3A2399221%2CTeletrac%2520Navman%3A24990%2CMTData%2520LLC%3A2629625%2CMTData%2520NZ%3A2632532%2CAutomation%2520Systems%2520%2526%2520Controls%2520Pty%2520Ltd%3A3066360%2CLionel%2520Samson%2520Sadleirs%2520Group%3A3114048%2CPeacock%2520Bros.%3A3219327%2CR%2526R%2520Corporate%2520Health%3A3276421%2CCombilift%2520Depot%3A3314927%2CConquest%2520Equipment%3A3361929%2CFischer%2520Plastic%2520Products%2520Pty%2520Ltd%3A3364182%2CMicrolise%3A33839%2CDepartment%2520of%2520Transport%3A349778%2CEnvirofluid%3A3552049%2CAllotrac%3A3575424%2CASCI%2520-%2520Australasian%2520Supply%2520Chain%2520Institute%3A3682349%2CTruckworld.com.au%3A3692277%2CAustralian%2520Sweeper%2520Company%3A3738673%2CTeletrac%2520Navman%2520%257C%2520Australia%3A3742855%2CFindex%3A3967749%2CKAB%2520Seating%2520Limited%3A4027652%2CTransficient%2520-%2520representing%2520TK'Blue%2520Agency%2520%2526%2520Global%2520Climate%2520Initiatives%2520in%2520Australia%3A42697837%2CANL%3A43609%2CLeopard%2520Systems%3A486727%2CULTIMATE%2520LED%3A48964277%2CGoodyear%2520%2526%2520Dunlop%2520Tyres%2520(Australia)%2520Pty%2520Ltd%2520Australia%3A53770224%2CLOAD%2520SMART%2520TRUCKING%2520INC%3A53829130%2CTransport%2520Certification%2520Australia%3A545896%2CTransurban%2520Group%3A55416951%2CDematic%2520Retrotech%3A56675%2CScaco%2520Pty%2520Ltd%3A58555205%2CGilbarco%2520Australia%2520Holding%2520Pty%2520Ltd%3A58681899%2CCincom%2520Systems%3A6160%2CSadleirs%2520Transport%3A62647764%2CCotewell%3A6423068%2CPort%2520of%2520Newcastle%3A6432433%2CNational%2520Heavy%2520Vehicle%2520Regulator%3A6520823%2CASAFE%2520AUSTRALASIA%2520Pty%2520Ltd%3A6590925%2CX-Pak%2520Global%3A6626434%2CDatanet%2520Asia%2520Pacific%2520Pty%2520Ltd%3A662994%2CUltimate%2520LED%2520Lights%3A67776597%2CWISTA%2520International%2520%3A68796735%2CWIM%2520Technologies%3A69544956%2CANC%2520Distribution%2520%3A71013196%2CTele%2520Radio%2520Group%2520Australia%2520%3A71622572%2CColliers%2520Engineering%2520%2526%2520Design%3A71630305%2CAI%2520Drive%2520By%2520GoLive%2520Monitoring%3A71992449%2CKBS%252C%2520a%2520Premier%2520Tech%2520Company%3A749880%2CNetstar%2520Australia%3A757162%2CTriTech%2520Lubricants%2520Australasia%3A7616600%2CEonmetall%2520Official%3A77077646%2CMuscat%2520Trailers%3A7750950%2CCombilift%2520Australia%3A7779401%2CBustle%2520Tech%3A7791457%2CM2MONEY%2520ONLINE%2520SERVICE%3A78910980%2CGilbarco%2520Veeder%2520Root%2520SEA%3A79842426%2CAUSTRALIAN%2520SWEEPER%2520CORPORATION%2520PTY%2520LTD%3A8274156%2CBailey%2520Ladders%3A9125623%2CGlobal%2520GPS%2520Tracking%3A9580664%2CGILBARCO%3A9739658&companyTimeScope=CURRENT&doFetchHeroCard=false&geoIncluded=101452733&logHistory=true&rsLogId=1322161036&searchSessionId=tkTbQPkmTXiU807BqoHmDg%3D%3D&selectedFilter=GE&seniorityIncluded=7%2C6%2C8&titleIncluded=Managing%2520Director%3A16%2CGeneral%2520Manager%3A17&titleTimeScope=CURRENT"

contract_chooser = "https://www.linkedin.com/sales/contract-chooser"




class Profile:

    def __init__(self, company, fullname, job_title, location, url):
        self._company = company
        self._job_title = job_title
        self._location = location
        self._url = url
        self._full_name = self.get_full_name(fullname)
        self._first_name = self.get_first_name(fullname)
        self._last_name = self.get_last_name(fullname)

    def get_company(self):
        return self._company
    
    def get_fullname(self):
        return self._fullname

    def get_job_title(self):
        return self._job_title

    def get_location(self):
        return self._location

    def get_url(self):
        return self._url

    def get_full_name(self, fullname):
        pattern = re.compile('[\W_]+', re.UNICODE)
        return re.sub(pattern, ' ', fullname)

    def get_first_name(self, fullname):
        name = HumanName(self.get_full_name(fullname))
        return name.first
    
    def get_last_name(self, fullname):
        name = HumanName(self.get_full_name(fullname))
        return name.last


def main():
    test_search()
    
    

def test_search():
    #Take record of time that this program started running.
    start_time = time.time()

    driver = webdriver.Chrome()

    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(driver)

    #Open an empty search page in Sales Navigator    
    #start_empty_search_in_sales_nav(driver)

    #Following line of code has been commented out because the Linkedin returns too many request message.
    #select_companies_in_search(driver, company_list)

    #Select CXO as a seniority level. 
    #select_seniority_in_search(driver, "CXO")
    #select_seniority_in_search(driver, "VP")
    #select_seniority_in_search(driver, "Director")

    #Search and then select Australia as geographical location of the leads.
    #search_geography_in_search(driver, "Australia")
    #select_geography_in_search(driver)
    #search_geography_in_search(driver, "New Zealand")
    #select_geography_in_search(driver)

    #Search and then select function of the leads.
    #search_function_in_search(driver, "Marketing")
    #select_function_in_search(driver)

    #Search and then select retail as industry of the leads.
    #search_industry_in_search(driver, "retail")
    #select_industry_in_search(driver)

    #Select Chief Marketing Officer as a title in search.
    #select_title_in_search(driver, 'managing director')
    #select_title_in_search(driver, 'general manager')

    driver.get(saved_search)
    
    
    #Zoom the browser to 60%.
    driver.execute_script("document.body.style.zoom='60%'")

    scroll_down(driver)

    #Get the number of pages in the search results. page_num is a string.
    page_num = get_num_of_search_result_pages(driver)
    print("You are now ready to move on to working with " + page_num + " pages.")
    
    #Get the number of search results in the current page.
    results_num = get_num_of_search_results_in_current_page(driver)
    print("You are now ready to move on to working with " + str(results_num) + " results in current page.")

    #Open each pages in the search. Append all page urls to page_urls list. 
    iterate_through_pages(driver)

    #Close the browser and its process to prevent out of memory issue.
    driver.quit()
    

    #Open each page in the search one by one. 
    for url in page_urls:
        #Open all results in current url one by one. Grab details and append it to leads list.
        temp_search(url)

    print("All results have been printed.")

    #Write the copied details into an excel file.
    write_leads_to_excel_file("supply.xlsx", "directors")
    print("All leads data have been written to xlsx file.")
    
    time.sleep(5)
    

    print("---This program took %s seconds ---" % (time.time() - start_time))    


def temp_search(url):

    #Make a temporary browser.
    temp_driver = webdriver.Chrome()

    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(temp_driver)

    #Bring current page
    temp_driver.get(url)
    
    #Go through all the search results in this page.
    iterate_through_results(temp_driver)
    #Close the browser and its process in the background.
    temp_driver.quit()



def log_into_linked_in_sales_nav(driver):    
    
    driver.get("https://www.linkedin.com")

    try:
        login_form_pw = driver.find_element_by_id('session_password')
        login_form_id = driver.find_element_by_id('session_key')
        login_form_btn = driver.find_element_by_class_name("sign-in-form__submit-button")
        
        file_id = open('file_id.txt','r')
        id = file_id.read()
        file_id.close()

        file_password = open('file_password.txt','r')
        password = file_password.read()
        file_password.close()

        login_form_id.send_keys(id)
        login_form_pw.send_keys(password)
        login_form_btn.send_keys(Keys.RETURN)

        driver.get("https://www.linkedin.com/sales/homepage")
        
    except StaleElementReferenceException:
        driver.refresh()
        log_into_linked_in_sales_nav(driver)



def start_empty_search_in_sales_nav(driver):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.ID,'global-typeahead-search-input')))
    
        search_bar = driver.find_element_by_id('global-typeahead-search-input');
        search_bar.send_keys(Keys.RETURN)

    except StaleElementReferenceException:
        driver.refresh()
        start_empty_search_in_sales_nav(driver)
        


def select_seniority_in_search(driver, level):

    try:
        wait = WebDriverWait(driver, 10)
        # Wait until seniority tab element is found.
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[10]')))
        # Make seniority variable refer to Seniority level tab element in the Sales Navigator.
        seniority = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[10]')
        #Click open the Seniority level tab.
        seniority.click()

        lv = '"'+level+'"'

        # XPATH must be as follows. The . checks the whole string value of the button element
        # Explanation is at https://stackoverflow.com/questions/23676537/xpath-for-button-having-text-as-new
        # '//form[@class="search-filter__form"]/ul/li[10]/div/div/div/ol/li/button[contains(.,' + lv + ' )]'
        
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[10]/div/div/div/ol/li/button[contains(.,' + lv + ' )]')))
        # Get seniority list
        seniority = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[10]/div/div/div/ol/li/button[contains(.,' + lv + ' )]')
        seniority.click()

    except Exception as e:
        print(e)
        driver.quit()




def select_function_in_search(driver, category):

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[11]')))
    
        function = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]')
        function.click()

        function_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/input')
        function_search_bar.send_keys(category)
        function_search_bar.send_keys(Keys.RETURN)
    except StaleElementReferenceException:
        driver.refresh()
        select_function_in_search(driver, category)


    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        function_country_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')
        function_country_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_function_in_search(driver, category)




def search_industry_in_search(driver, industry):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[8]')))
        industry_filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[8]')
        industry_filter.click()
        industry_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[8]//div[@class="ph4 pb4"]/input')
        industry_search_bar.send_keys(industry)
        industry_search_bar.send_keys(Keys.RETURN)
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        search_industry_in_search(driver, industry)


def select_industry_in_search(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[8]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        industry_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[8]//div[@class="ph4 pb4"]/ol/li[1]/button')
        industry_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_industry_in_search(driver)




def search_function_in_search(driver, function):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[11]')))
        function_filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]')
        function_filter.click()
        function_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/input')
        function_search_bar.send_keys(function)
        function_search_bar.send_keys(Keys.RETURN)
        
    except (StaleElementReferenceException , TimeoutException) as e:
        driver.refresh()
        print(e)
        search_function_in_search(driver, function)
        


def select_function_in_search(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        function_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[11]//div[@class="ph4 pb4"]/ol/li[1]/button')
        function_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException) as e:
        driver.refresh()
        print(e)
        select_function_in_search(driver)




def search_geography_in_search(driver, country):

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]')))
        geography = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]')
        geography.click()
        geography_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/input')
        geography_search_bar.send_keys(country)
        geography_search_bar.send_keys(Keys.RETURN)
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        search_geography_in_search(driver, country)


def select_geography_in_search(driver):    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')))
    
        geography_country_btn = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[5]//div[@class="ph4 pb4"]/ol/li[1]/button')
        geography_country_btn.send_keys(Keys.RETURN)

    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_geography_in_search(driver)


def select_title_in_search(driver, title):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[12]')))
    
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]//div[@class="ph4 pb4"]/input')
        filter_search_bar.send_keys(title)
        filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        select_title_in_search(driver, title)
        driver.refresh()
        

def select_titles_in_search(driver, titles):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[12]')))
    
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[12]//div[@class="ph4 pb4"]/input')
        for title in titles:
            filter_search_bar.send_keys(title)
            filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_titles_in_search(driver, titles)


def select_companies_in_search(driver, companies):
    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[7]')))
    
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]//div[@class="ph4 pb4"]/input')
        for company in companies:
            filter_search_bar.send_keys(company)
            filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_companies_in_search(driver, companies)

def select_a_company_in_search(driver, company):    
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//form[@class="search-filter__form"]/ul/li[7]')))
    
        filter = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]')
        filter.click()

        filter_search_bar = driver.find_element(By.XPATH, '//form[@class="search-filter__form"]/ul/li[7]//div[@class="ph4 pb4"]/input')
        filter_search_bar.send_keys(company)
        filter_search_bar.send_keys(Keys.RETURN)
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        select_companies_in_search(driver, companies)

        
def get_num_of_search_result_pages(driver):
# I want to know the number of pages of search results   

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]')))
    
        page_num = driver.find_element(By.XPATH, '//section[@id="results"]/div/nav/ol[@class="search-results__pagination-list"]/li[last()]/button').text

    except NoSuchElementException:
        print("There is 1 page.")
        return "1"
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        get_num_of_search_result_pages(driver)

    return page_num

def get_num_of_search_results_in_current_page(driver):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]')))
        html_list = driver.find_elements(By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li')
        results_num = len(html_list)
    except StaleElementReferenceException:
        results_num = 0
    except TimeoutException:
        driver.refresh() 
        get_num_of_search_results_in_current_page(driver)
        
    return results_num


def iterate_through_pages(driver):
    
    page_num = int(get_num_of_search_result_pages(driver))
    curr = 1

    #Start populating a list of all page urls, starting with current page url.
    page_urls.append(driver.current_url)

    #iterate_through_results(driver)
    while curr < page_num:
        curr+=1
        
        try:
            wait = WebDriverWait(driver, 10)
            element = wait.until(EC.presence_of_element_located((By.XPATH, '//div/nav/ol[@class="search-results__pagination-list"]')))
            nextPage = driver.find_element(By.XPATH, '//div/nav/ol[@class="search-results__pagination-list"]/li[@class="selected cursor-pointer"]/following-sibling::li/button')
            nextPage.send_keys(Keys.RETURN)
            
            time.sleep(2)
            page_urls.append(driver.current_url)
        except (StaleElementReferenceException , TimeoutException):
            curr-=1
            driver.refresh()


def iterate_through_results(driver):
    results_num = get_num_of_search_results_in_current_page(driver)
    scroll_down(driver)

    if results_num > 0:
        curr = 1
        while curr <= results_num:
            #open_search_results(driver, curr)
            get_profile_data_from_search_result(driver, curr)
            curr+=1
    else:
        #Do nothing.
        curr = 0


# I want to go through the results in the page and print one by one on console.
# This function assumes that the driver is currently at a sales navigator page with results showing from search.
# For each result, this function grab profile data: full name, first name, last name, location, position, company and url of LinkedIn profile.
# This populates leads list variable so it can be written to an excel file.
def get_profile_data_from_search_result(driver, pointer):
    
    #Get the number of results in the current page.
    results_num = get_num_of_search_results_in_current_page(driver)

    #Initialise this WebDriverWait instance so I can use in the loop below.
    wait = WebDriverWait(driver, 10)

    #Use string of pointer for XPATH
    pointer_str = str(pointer)

    try:
        #Wait until full name appears in DOM.
        elem1 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dt/a')))
        #Get the full name.
        fullname = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dt/a').text
        url = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dt/a').get_attribute('href')

    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
        
    try:
        #Wait until position appears in DOM.
        elem2 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[2]/span[1]')))
        #Get the position. 
        job_title = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[2]/span[1]').text

    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)
        
    try:                               
        #Wait until position appears in DOM.
        elem3 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[2]/span[2]/span/a/span[1]')))
        #Get the position. 
        company = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[2]/span[2]/span/a/span[1]').text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)

    try:       
        #Wait until position appears in DOM.
        elem4 = wait.until(EC.presence_of_element_located((By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[4]/ul/li')))
        #Get the position. 
        location = driver.find_element(By.XPATH, '//ol[@class="search-results__result-list"]/li[' + pointer_str + ']/div[2]/div/div/div/article/section[1]/div[1]/div/dl/dd[4]/ul/li').text
    except StaleElementReferenceException:
        driver.refresh()
        scroll_down(driver)
        get_profile_data_from_search_result(driver, pointer)

    person = Profile(company, fullname, job_title, location, url)
    leads.append(person)

    print(person._company + "%^&" + person._full_name + "%^&" + person._first_name + "%^&" + person._last_name + "%^&" + person._location + "%^&" + person._job_title + "%^&" + person._url)
    print("  ")



def open_search_results(driver, curr):
    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li['+ str(curr) + ']/div[2]/div/div/div/article/section[@class="result-lockup"]/div/div/dl/dt[@class="result-lockup__name"]/a')))

        url = driver.find_element(By.XPATH, '//section[@id="results"]/div/div/ol[@class="search-results__result-list"]/li['+ str(curr) + ']/div[2]/div/div/div/article/section[@class="result-lockup"]/div/div/dl/dt[@class="result-lockup__name"]/a').get_attribute('href')
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[1])
        driver.get(url);
        
        grab_details(driver)

        time.sleep(2)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
        
    except (StaleElementReferenceException , TimeoutException):
        driver.refresh()
        open_search_results(driver, curr)


def grab_details(driver):

    driver.execute_script("document.body.style.zoom='60%'")

    try:
        wait = WebDriverWait(driver, 10)
        
        elem_fullname = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dt/span')))
        fullname = driver.find_element(By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dt/span').text

        #This if statement handles a case where clicking a link brings up locked Linkedin profile.
        if fullname != "LinkedIn Member":
        
            elem_location = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dd[@class="mt4 mb0"]/div')))
            elem_position = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dt')))
            elem_company = wait.until(EC.presence_of_element_located((By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dd[1]/span[2]')))
            
            job_title = driver.find_element(By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dt').text
            
            location = driver.find_element(By.XPATH, '//div[@class="container"]/div/div/div/div/dl/dd[@class="mt4 mb0"]/div').text     
            company = driver.find_element(By.XPATH, '//section[@id="profile-positions"]/div/ul/li[1]/dl/dd[1]/span[2]').text  
            url = driver.current_url

            person = Profile(company, fullname, job_title, location, url)
            leads.append(person)

            print(person._company + "%^&" + person._full_name + "%^&" + person._first_name + "%^&" + person._last_name + "%^&" + person._location + "%^&" + person._job_title + "%^&" + person._url)
            print("  ")
    
        else:
            print(driver.current_url)
            print("This is a locked Linkedin Member in Sales Navigator.")

    except StaleElementReferenceException:
        driver.refresh()
        grab_details(driver)


    

def scroll_down(driver):
    
    height = driver.execute_script("return document.documentElement.scrollHeight")

    #Divide the scroll height into 6 equal sections so driver can stop at each section.
    sect1 = height/6
    sect2 = sect1 * 2
    sect3 = sect1 * 3
    sect4 = sect1 * 4
    sect5 = sect1 * 5

    #Stop at each sections ending with the bottom of the page. Make sure all results load on the page.    
    driver.execute_script("window.scrollTo(0, " + str(sect1) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(sect2) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(sect3) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(sect4) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(sect5) + ");")
    time.sleep(1)
    driver.execute_script("window.scrollTo(0, " + str(height) + ");")
    time.sleep(1)

def write_leads_to_excel_file(file_name, sheet_name):
    # file_name e.g "leads_.xlsx"
    workbook = xlsxwriter.Workbook(file_name)
    # sheet_name e.g "HCL Appscan 2022"
    worksheet = workbook.add_worksheet(sheet_name)

    row = 1
    col = 0

    header = ["COMPANY", "FULL NAME", "FIRST NAME", "LAST NAME", "JOB TITLE", "LOCATION", "LINKEDIN URL"]

    for hd in header:
        worksheet.write(0, col, hd)
        col+=1

    for lead in leads:
        worksheet.write(row, 0, lead._company)
        worksheet.write(row, 1, lead._full_name)
        worksheet.write(row, 2, lead._first_name)
        worksheet.write(row, 3, lead._last_name)
        worksheet.write(row, 4, lead._job_title)
        worksheet.write(row, 5, lead._location)
        worksheet.write(row, 6, lead._url)
        print(str(row) + " leads written to file.")
        row+=1        
    
    workbook.close()

main()
