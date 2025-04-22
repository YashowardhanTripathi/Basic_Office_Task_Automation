import requests
import json
import os

def goAnywhere_tool() :
    # API details
    base_url = "https://mftweb.domain.com:8001/goanywhere/rest/gacmd/v1/projects"
    username = os.getenv("GA_USERNAME")
    password = os.getenv("GA_PASSWORD")

    print(f"The username is {username} and password is {password}")
    # JSON payload
    payload = {

        "runParameters": {
            "project": "Project Path",
            "domain": "domain-ecm",
            "jobName": "Job Name",
            "jobQueue": "jobs-queue-ecm",
            "mode": "batch",
            "priority": "5"
        }

    }

    # Headers
    headers = {
        "Content-Type": "application/json"
    }

    # Making the POST request
    response = requests.post(base_url, auth=(username, password), headers=headers, json=payload)

    # Checking the response
    if response.status_code == 200:
        print("Success:", response.json())
    else:
        print("Error:", response.status_code, response.text)

def main() :
    goAnywhere_tool()
if __name__ ==  "__main__":
    main()