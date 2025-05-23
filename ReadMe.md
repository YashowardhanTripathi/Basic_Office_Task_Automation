Basic Daily TaskAutomate -- Application Documentation
=====================================================

**Basic Daily Task Automate -Application Documentation**
--------------------------------------------------------

### **1\. Introduction**

The Basic Daily Task Automate application is designed to reduce manualintervention in performing common daily tasks, specifically those related touser access management, Managed File Transfer (MFT) error handling, andGoAnywhere error alert resolution. This application aims to improve efficiency,reduce errors, and free up personnel for more strategic work.

The application is currently delivered as full source code, with plans toprovide a packaged executable (.exe) file in future releases. A web applicationinterface is also available, offering a user-friendly way to access the tool'sfunctionalities.

### **2\. Application Overview**

The Basic Daily Task Automate application is available in three forms:

·        **.exe File:** A standaloneexecutable for easy deployment and use.

·        **Web Application:** A web-basedinterface providing access to specific functions.

·        **Full Source Code:** The completesource code, offering maximum flexibility and customization.

### **3\. Key Features and Use Cases**

The application automates the following key tasks:

·        **User Access Management:**Automates the process of providing or terminating user access to systems andapplications. This can include creating accounts, assigning permissions, andrevoking access.

·        **MFT Error Handling:** Automatesthe resolution of errors encountered during Managed File Transfer processes.This may involve retrying transfers, notifying support teams, or takingcorrective actions.

·        **GoAnywhere Error Alert Resolution:**Automates the resolution of errors and alerts generated by the GoAnywhere MFTplatform. This can include investigating error logs, applying fixes, andescalating critical issues.

**Use Cases:**

·        Reduce the time and effort required to provisionor de-provision user access.

·        Minimize disruptions caused by MFT errors byenabling rapid automated resolution.

·        Improve the efficiency of GoAnywhere operationsby automating the handling of routine errors.

·        Decrease the risk of human error in performingrepetitive daily tasks.

·        Provide a consistent and auditable way toperform critical tasks.

### **4\. Application Architecture**

The application is built using the following technologies:

·        **Python:** The core logic andautomation scripts are written in Python.

·        **Flask:** The web applicationinterface is built using the Flask web framework.

·        **HTML:** The front-end of the webapplication is built using HTML.

·        **JavaScript:** The front-end ofthe web application uses JavaScript for interactivity.

### **5\. Setup and Installation**

The setup and installation process varies depending on the application form:

#### **5.1 Full Source Code**

1.     **Prerequisites:**

·        Python 3.x installed.

·        Required Python packages (install via pipinstall -r requirements.txt, if applicable).

·        GoAnywhere credentials.

2.     **Installation:**

·        Download or clone the full source code.

·        Navigate to the source code directory.

·        Install the necessary Python packages (if arequirements.txt file is provided).

·        Configure GoAnywhere credentials:

·        Edit the Goanywhere.py file.

·        Set the username and password variables to thecorrect values.

3.     **Running the Application:**

·        Open a terminal or command prompt.

·        Navigate to the source code directory.

·        Run the application using the command: pythonBasic\_daily.py

#### **5.2 .exe File**

1.     **Prerequisites:**

·        GoAnywhere credentials.

2.     **Installation:**

·        Copy the .exe file to the desired location.

3.     **Configuration:**

·        Open a command prompt.

·        Set the GoAnywhere credentials as environmentvariables (one-time setup):

·        setx USERNAME

·        setx PASSWORD

·        **Note:** Replace and with theactual credentials. These are system-wide environment variables.

4.     **Running the Application:**

·        Double-click the .exe file to run theapplication.

#### **5.3 Web Application**

1.     **Prerequisites:**

·        Web server to host the application (e.g.,Apache, Nginx).

·        Python and Flask (for the backend).

·        GoAnywhere credentials.

2.     **Installation:**

·        Deploy the application files to the web server.

·        Configure the web server to run the Flaskapplication.

·        Configure GoAnywhere credentials (similar to theFull Source Code instructions, depending on the backend implementation).

3.     **Running the Application:**

·        Access the application through a web browserusing the appropriate URL.

·        The specific functions will be visible in theweb interface, and you can choose and run the specific function.

### **6\. Future Enhancements**

The following enhancements are planned for future versions of theapplication:

·        **AI-Powered Automation:**Implement machine learning models (potentially using TensorFlow) to enable theapplication to learn from past tasks and automate them more intelligently. Thiscould involve predicting task requirements, optimizing workflows, andproactively addressing potential issues.

·        **ServiceNow Integration:**Integrate with the ServiceNow API to:

·        Automatically update ticket statuses based ontask execution results.

·        Create new tickets for failed tasks orexceptions.

·        **Enhanced Alerting:** Implementmessage alerts for P1 and P2 incidents to prevent SLA breaches. This willinvolve integrating with a messaging platform (e.g., Slack, Microsoft Teams) toprovide timely notifications of critical issues.

·        **24/7 Availability**: Deploy theapplication to a dedicated server to ensure continuous operation andavailability.

### **7\. Challenges**

The following challenges need to be addressed:

·        **ServiceNow API Access:** Lack ofaccess to the ServiceNow API is currently preventing full integration with theticketing system.

·        **Dedicated Server:** Theapplication requires a dedicated server with 24/7 availability to ensurecontinuous operation, especially for automated tasks and alerts.

Additional Note:
----------------

·        The application is designed to be scalable andadaptable to future requirements.

·        For further customization or troubleshooting,refer to the source code documentation or contact the development team.

·        If you require additional details or havespecific questions, please reach out to the team for clarification.