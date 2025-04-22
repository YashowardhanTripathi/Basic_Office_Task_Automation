from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException, WebDriverException
import time

def AlevateAccessSelenium(email, name,Auths):

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
    input_element.send_keys('password')
    input_element.send_keys(Keys.RETURN)
    input_element.click()
    time.sleep(6)
    input_element = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[2]/div/div/div/div[2]/div/button')
    input_element.click()
    time.sleep(6)
    # Form Filling
    # User ID
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[1]/input')
    input_element.send_keys(name)
    input_element.send_keys(Keys.RETURN)
    time.sleep(3)
    # Type
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[2]/input')
    input_element.send_keys('user')
    input_element.send_keys(Keys.RETURN)
    time.sleep(3)
    # First Name
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[3]/input')
    firstname = name.split(".")
    input_element.send_keys(firstname[0])
    input_element.send_keys(Keys.RETURN)
    time.sleep(3)
    # Last Name
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[4]/input')
    lastname  = name.split(".")
    input_element.send_keys(lastname[1])
    input_element.send_keys(Keys.RETURN)
    time.sleep(3)
    # Email
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[8]/input')
    input_element.send_keys(email)
    input_element.send_keys(Keys.RETURN)
    time.sleep(3)
    #Date Format
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[9]/select')
    input_element.send_keys(Keys.RETURN)
    input_element.click()
    time.sleep(3)
    # Preffered language
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[15]/input')
    input_element.send_keys('en')
    input_element.send_keys(Keys.RETURN)
    time.sleep(3)
    # User type
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[18]/input')
    input_element.send_keys('regular')
    input_element.send_keys(Keys.RETURN)
    time.sleep(3)
    # Authorization
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[19]/textarea')
    input_element.send_keys(Auths)
    input_element.send_keys(Keys.RETURN)
    time.sleep(3)
    # Approver Level
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[2]/div[21]/textarea')
    input_element.send_keys('[]')
    input_element.send_keys(Keys.RETURN)
    # Add User Button click
    input_element = driver.find_element(By.XPATH, '/html/body/div[2]/div/div/div[3]/button[2]')
    input_element.send_keys(Keys.RETURN)
    time.sleep(4)

    try:
        corecontent = driver.find_element(By.TAG_NAME, "body").text
        print(corecontent)  # Print extracted text if successful
    except NoSuchElementException:
        corecontent = "Error: Element with TAG_NAME 'body' not found."
    except WebDriverException as e:
        corecontent = f"Selenium WebDriver Error: {e}"
    except Exception as e:
        corecontent =  f"Unexpected Error: {e}"

    print("Check checkkk:- ",corecontent)
    if "already" in corecontent:
        corecontent = "User already exist"
    else:
        corecontent = f"User {name} is created"
    return corecontent
    # time.sleep()
    driver.quit()

def main():
    AlevateAccessSelenium()
if __name__ == '__main__':
    main()