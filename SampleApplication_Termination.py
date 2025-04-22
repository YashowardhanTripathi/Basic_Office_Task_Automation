from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

# from Basic_Daily import TerminateUser
import time



def TerminateUser(Usermail=""):
    # Usermail = "abc.p@mckesson.com"
    Chrome_Option = Options()
    Chrome_Option.add_argument("--headless")
    Chrome_Option.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
    Chrome_Option.add_argument("--ignore-certificate-errors")
    Chrome_Option.add_argument("--allow-insecure-localhost")
    Chrome_Option.add_argument("--disable-web-security")
    Chrome_Option.add_argument("--disable-site-isolation-trials")


    driver = webdriver.Chrome(options=Chrome_Option)
    driver.get("Application URL")

    # Now, you can interact with the website without having to log in again



    # Alevate Passsword Mckesson123789

    input_element = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div/div/input[1]')
    input_element.send_keys('abc.abc.com')
    input_element.send_keys(Keys.RETURN)
    time.sleep(6)
    input_element = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div/div/button[2]')
    input_element.send_keys(Keys.RETURN)
    time.sleep(6)
    input_element = driver.find_element(By.XPATH, '/html/body/div/div[2]/div[2]/div/div/input[2]')
    input_element.send_keys('passowrd')
    input_element.send_keys(Keys.RETURN)
    input_element.click()
    time.sleep(6)
    input_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div/div[2]/div/form/div[1]/input')
    input_element.send_keys(Usermail)
    input_element.send_keys(Keys.RETURN)
    time.sleep(6)
    # Form Filling
    # User ID
    input_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div/div[2]/div/form/div[2]/button')
    input_element.click()
    time.sleep(3)

    page_text = driver.find_element(By.TAG_NAME, "body").text

    # You can submit the form, click buttons, etc.
    # submit_button = driver.find_element(By.NAME,"Submit")
    # submit_button.click()
    # driver.implicitly_wait(5)

    # response_message = driver.find_element(By.ID, "response")  # Replace with the actual ID or class
    print("Response message:", page_text)

    if "[]" in page_text:
        print(f"User: {Usermail} found")
        input_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div/div[2]/table/tbody/tr/td[17]/button[2]/i')
        input_element.click()
        # print (f"Yash Auths {auths[start:end].strip()}")
        return page_text
    else:
        print(f"User: {Usermail} not found")
        return f"User: {Usermail} Not Found!"




    driver.quit()
    return page_text

def main ():
    TerminateUser()

if __name__ == "__main__":
    main()

