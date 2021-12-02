from navigator import *

class Lusha:
    
    def __init__(self, email, mobile):
        self._email = email
        self._mobile = mobile

    def get_email(self):
        return self._email
    
    def get_mobile(self):
        return self._mobile

def try_lusha():
    
    #Load Lusha Extension to driver.
    extension='E:/python-selenium-code/10.3.2_0.crx'
    options = webdriver.ChromeOptions()
    options.add_extension(extension)
    driver = webdriver.Chrome(chrome_options=options)

    #Log into Linkedin Sales Navigator.
    log_into_linked_in_sales_nav(driver)

    #Log into Lusha then closes the browser tab.
    log_into_lusha(driver)

    #Wait until driver moves to far left first tab.
    time.sleep(2)
    #Refresh the tab because it takes refresh to update Lusha login.
    driver.refresh()
    #Give enough time for Lusha to appear on Linkedin page.
    time.sleep(2)

    #Open Lusha extension.
    open_lusha(driver)
    time.sleep(1)
    #Click to agree to Lusha privacy policy
    agree_to_lusha_privacy_policy(driver)
    time.sleep(1)

    #Get a Linkedin account profile.
    driver.get("https://www.linkedin.com/in/john-m-b44ab3108/")
    time.sleep(2)
    
    #Open Lusha extension on thia profile.
    open_lusha(driver)
    time.sleep(1)

    try:
        if check_lusha_details_exist(driver):
            print("Lusha has contact details for this one.")
        else:
            print("Lusha has no contact details for this one.")
    except Exception as e:
        print(e)

#This function is still in progress. The Lusha handles Bot by asking user to select images.
def log_into_lusha(driver):

    driver.switch_to.window(driver.window_handles[1])
    driver.get("https://auth.lusha.com/login")

    try:
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH, '//div[@id="__next"]/div/div/div/span/form/span[1]/div/input')))

        login_form_id = driver.find_element(By.XPATH, '//div[@id="__next"]/div/div/div/span/form/span[1]/div/input')
        login_form_pw = driver.find_element(By.XPATH, '//div[@id="__next"]/div/div/div/span/form/span[2]/div/input')
        login_form_btn = driver.find_element(By.XPATH, '//div[@id="__next"]/div/div/div/span/form/button')

        file_id = open('lusha_id.txt','r')
        id = file_id.read()
        file_id.close()

        file_password = open('lusha_password.txt','r')
        password = file_password.read()
        file_password.close()

        login_form_id.send_keys(id)
        login_form_pw.send_keys(password)
        login_form_btn.send_keys(Keys.RETURN)
        time.sleep(15)
        
    except StaleElementReferenceException:
        driver.refresh()
        log_into_lusha(driver)

    driver.close()
    driver.switch_to.window(driver.window_handles[0])


def get_lusha():

    #Let's check if the lusha found any details. Did it return details?
    if check_lusha_could_not_find(driver):
        #If lusha
        return Lusha("no email", "no mobile")

    
    if check_lusha_show_btn(driver):
        #True means the button needs a click to load the email and phone number.
        btn = driver.find_element(By.XPATH, '//div[@class="contact-header"]/div/div[@class="contact-action-container"]/div[@class="save-to-action-buttons"]/button')
        btn.click()

    # I need to implement these functions.
    #Check if there are any details
    if check_lusha_details(driver):
        #Check if there are any email addresses.
        if check_lusha_email(driver):
            email = get_lusha_email(driver)
        #Check if there are any phone numbers.
        if check_lusha_mobile(driver):
            mobile = get_lush_mobile(driver)
            

def agree_to_lusha_privacy_policy(driver):

    try:
        lusha_frame = driver.find_element_by_id("LU__extension_iframe")
        driver.switch_to.frame(lusha_frame)
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//div[@class="privacy-approval-buttons-container"]/button')))
        button = driver.find_element(By.XPATH,'//div[@class="privacy-approval-buttons-container"]/button')
        ActionChains(driver).move_to_element(button).click(button).perform()
        driver.switch_to.default_content()
    except  Exception as e:
        print(e)

def open_lusha(driver):

    try:
        wait = WebDriverWait(driver, 10)
        elem_fullname = wait.until(EC.element_to_be_clickable((By.ID, 'LU__extension_badge_main')))
        button = driver.find_element(By.ID, 'LU__extension_badge_main')
        ActionChains(driver).move_to_element(button).click(button).perform()
        
    except(StaleElementReferenceException , TimeoutException):
        driver.refresh()
        open_lusha(driver)


def check_lusha_details_exist(driver):

    #Try and check if the lusha has save contact button.
    try:
        lusha_frame = driver.find_element_by_id("LU__extension_iframe")
        driver.switch_to.frame(lusha_frame)
        wait = WebDriverWait(driver, 10)
        element = wait.until(EC.presence_of_element_located((By.XPATH,'//div[@class="save-to-actions-buttons"]/button')))
        btn = driver.find_element(By.XPATH,'//div[@class="save-to-actions-buttons"]/button')
        driver.switch_to.default_content()
        return True        
    except NoSuchElementException:
        driver.switch_to.default_content()
        return False
    
    return False



def check_lusha_details(driver):
    try:
        contact_details = driver.find_elements(By.XPATH, '//div[@class="contact-header"]/ul[@class="contact-details"]/li')
        return True
    except NoSuchElementException:
        return False

def check_lusha_could_not_find(driver):
    try:
        empty_state = driver.find_element(By.XPATH, '//div[@class="enrich-empty-state-text"]/div[1]')
        return True
    except NoSuchElementException:
        return False


def check_lusha_show_btn(driver):

    try:
        show_contact_btn_text = driver.find_element(By.XPATH, '//div[@class="contact-header"]/div/div[@class="contact-action-container"]/div[@class="save-to-action-buttons"]/button/span').text
    except NoSuchElementException:
        return False

    if show_contact_btn_text == "Show contact":
        return True
    
    return False
