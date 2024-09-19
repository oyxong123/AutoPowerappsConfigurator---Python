# Run script whenever computer wake up from sleep or powers on.
# Schedules register-attendance jobs based on config file, and runs missed jobs immediately.

import time 
import datetime
import os
import schedule
import pytz
import threading
import pystray
import sys
import wmi
from PIL import Image
from typing import List
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_auto_update.webdriver_auto_update import WebdriverAutoUpdate
from win10toast_persist import ToastNotifier

def LogInPowerappsAttendance(userEmail, userPassword) -> webdriver:
    # Autoupdate chromedriver.
    driverDirectory = "C:/Windows"
    # WebdriverAutoUpdate(driverDirectory).check_driver()  # removed due to not working correctly.
    
    # Configure driver options and services.
    options = webdriver.ChromeOptions()
    options.page_load_strategy = 'normal'
    options.add_argument("--headless")  # only apply in production.
    options.add_argument("--window-size=1920x1080")
    logPath = "C:/Users/Admin/Documents/Selenium/AutoPowerapps"
    service = webdriver.ChromeService(service_args=['--log-level=DEBUG', '--append-log', '--readable-timestamp'], log_output=logPath)
    driver = webdriver.Chrome(options, service)
    driver.implicitly_wait(20)
    
    # Access PowerApps UG Attendance website.
    driver.get("https://apps.powerapps.com/play/e/default-84187be3-037e-41ec-889c-a150fe476432/a/afab9b41-ef46-4e5d-988b-2d0dce08234d?tenantId=84187be3-037e-41ec-889c-a150fe476432&sourcetime=1705632106028&source=portal")
    time.sleep(5)
    
    # Handle Email Login.
    elementInputEmail = driver.find_element(By.NAME, "loginfmt")
    elementInputEmail.send_keys(userEmail)
    elementSubmit = driver.find_element(By.CSS_SELECTOR, "[type = submit]")
    # elementSubmit.click()  #no idea why button auto clicked.
    time.sleep(5)
    
    # Handle Password Login.
    elementInputPassword = driver.find_element(By.NAME, "passwd")
    elementInputPassword.send_keys(userPassword)
    elementSubmit = driver.find_element(By.CSS_SELECTOR, "[type = submit]")
    # elementSubmit.click()  #no idea why button auto clicked.
    time.sleep(3)
    
    # Handle Login Confirmation.
    elementSubmit = driver.find_element(By.CSS_SELECTOR, "[type = submit]")
    elementSubmit.click()
    time.sleep(10)
    
    # Special IFrame locator required to interact with PowerApps Web Player.
    driver.switch_to.frame("fullscreen-app-host")
    time.sleep(5)
    
    # Return webdriver object.
    return driver

def RegisterAttendance(email, password, course):
    try:
        # Set script initialize state.
        taskStatus = True
        
        # Notify computer squaring has started for what category.
        notif = ToastNotifier()
        notif.show_toast("Auto Powerapps", f"Subject: {course} \nState: Registering", duration=5)
        
        # Log in to PowerApps UG Attendnace App.
        driver = LogInPowerappsAttendance(email, password)
        
        # Handle New Attendance Creation.
        elementNewAttendance = driver.find_element(By.CSS_SELECTOR, "[data-control-name = IconNewItem1]")
        elementNewAttendance.click()
        time.sleep(2)
        
        # Find specific course based on config file.
        elementInputCourse = driver.find_element(By.TAG_NAME, "input")
        elementInputCourse.send_keys(course)
        time.sleep(2)
        
        # Select the specific course searched.
        elementCount = len(driver.find_elements(By.CSS_SELECTOR, "[data-control-name=NextArrow3]"))
        if (elementCount <= 0):
            raise Exception("No specified course found.")
        elif (elementCount > 1):
            raise Exception("Multiple course selections found.")
        elementOptionCourse = driver.find_element(By.CSS_SELECTOR, "[data-control-name=NextArrow3]")
        elementOptionCourse.click()
        time.sleep(2)
        
        # Register specified course's attendance.
        elementConfirmCourse = driver.find_element(By.CSS_SELECTOR, "[data-control-name=IconAccept1]")
        elementConfirmCourse.click()
        time.sleep(2)
        
        # Confirm that the specified course has already been registered.
        courseRegisteredList = []
        elementCount = len(driver.find_elements(By.CSS_SELECTOR, "[data-control-name = Subtitle1]"))
        if elementCount > 0:
            elementCourseRegisteredList = driver.find_elements(By.XPATH, "//div[@data-control-name = 'Subtitle1']/div[1]/div[1]/div[1]/div[1]")
            for element in elementCourseRegisteredList:
                courseRegisteredList.append(element.text)
        if course not in courseRegisteredList:
            raise Exception("Registered attendance not found.")
        time.sleep(5)
        
        # End selenium session.
        driver.quit()
        
        # Update the local registerred attendances log.
        fileRegisteredAttendances = open(pathRegisteredAttendnaces, "a")
        fileRegisteredAttendances.write(f"{course}\n")
        fileRegisteredAttendances.close()
        
    except Exception as exception:
        # Catch exception.
        taskStatus = False
        errorMessage = str(exception)
        driver.quit()
        
    # Add a notification to notify me about the success or failure.
    notif = ToastNotifier()
    if taskStatus:
        notif.show_toast("Auto Powerapps", f"Subject: {course}\nState: Registered", duration=5)
        print(f"\nAuto Powerapps script completed. Operation Success.\nSubject: {course}\nState: Registered\n") 
    else:
        notif.show_toast("Auto Powerapps", f"Subject: {course}\nState: Failed\n" + errorMessage, duration=5)
        print(f"\nAuto Powerapps script failed. Operation failed.\nSubject: {course}\nState: Failed\n" + errorMessage) 
    print("Scheduler Continue Running.\n")
    
def RetrieveRegisteredAttendances(email, password) -> List[str]:
    # Log in to PowerApps UG Attendnace App.
    driver = LogInPowerappsAttendance(email, password)
    
    # Retrieve registered courses.
    courseRegisteredList = []
    elementCount = len(driver.find_elements(By.CSS_SELECTOR, "[data-control-name = Subtitle1]"))
    if elementCount > 0:
        elementCourseRegisteredList = driver.find_elements(By.XPATH, "//div[@data-control-name = 'Subtitle1']/div[1]/div[1]/div[1]/div[1]")
        for element in elementCourseRegisteredList:
            courseRegisteredList.append(element.text)
    
    # End selenium session.
    driver.quit()

    # Return all registered courses.
    return courseRegisteredList

def OpenConfigFile():
    # Forced type-unface way of passing argument. Make sure the correct variable "pathConfig" is allocated in runtime. 
    os.startfile(pathConfig)
    print("\nConfig file opened.\n")

def ReloadConfigFile():
    os.execv(sys.executable, ['python3'] + sys.argv)  #needupdate

def OpenRunHistory():
    # Forced type-unface way of passing argument. Make sure the correct variable "pathRunHistory" is allocated in runtime. 
    os.startfile(pathRunHistory)
    print("\nRun history opened.\n")
    
def QuitScript(sysTray):
    properExitCode = 50
    try:
        print("\nExiting script.\n")
        sysTray.stop()
        sys.exit(properExitCode)
    except SystemExit as e:
        if e.code != properExitCode:
            raise
        else:
            os._exit(properExitCode)
    
def BuildSystemTray(imgPath) -> pystray.Icon:
    imgTray = Image.open(imgPath)
    sysTray = pystray.Icon("AutoPowerappsConfigurator", imgTray, "Auto Powerapps Configurator")
    item1 = pystray.MenuItem("Edit Config", OpenConfigFile) 
    item2 = pystray.MenuItem("Reload Config", ReloadConfigFile) 
    item3 = pystray.MenuItem("Run History", OpenRunHistory) 
    item4 = pystray.MenuItem("Exit", lambda: QuitScript(sysTray))
    menuTray = pystray.Menu(item1, item2, item3, item4)
    sysTray.menu = menuTray
    sysTray.run_detached()
    return sysTray

def HandleWindowsLatestEvent(winEvent, sysTray):
    # Reference: https://learn.microsoft.com/en-us/windows/win32/power/wm-powerbroadcast
    if winEvent.EventType == 10:  # PBT_APMPOWERSTATUSCHANGE, When power status of computer changes.
        pass 
    if winEvent.EventType == 18:  # PBT_APMRESUMEAUTOMATIC, When computer resume from low-powered state (eg. waking up from sleep). Compared to PBT_APMRESUMESUSPEND, triggers first if an input signal is sent (eg. power button is pressed).
        QuitScript(sysTray)
    if winEvent.EventType == 7:  # PBT_APMRESUMESUSPEND, When computer resume from low-powered state (eg. waking up from sleep). 
        QuitScript(sysTray)
    if winEvent.EventType == 4:  # PBT_APMSUSPEND, When computer is suspending operation (eg. entering sleep).
        QuitScript(sysTray)
    if winEvent.EventType == 32787:  # PBT_POWERSETTINGCHANGE, When a change in computer power setting occurs (eg. changing power supply from battery to AC power).
        pass 

loggedInUser = os.getlogin()
pathSeleniumScript = f"C:/Users/{loggedInUser}/Documents/VSCodeProjects/Python"
pathConfig = f"C:/Users/{loggedInUser}/Documents/AutoPowerapps/config.txt"
pathConfigTemplate = f"C:/Users/{loggedInUser}/Documents/AutoPowerapps/config_Template.txt"
pathRunHistory = f"C:/Users/{loggedInUser}/Documents/AutoPowerapps/run_history.txt"
pathRegisteredAttendnaces = f"C:/Users/{loggedInUser}/Documents/AutoPowerapps/run_registered-attendances.txt"
imgPath = f"C:/Users/{loggedInUser}/Documents/AutoPowerapps/autopowerapps-icon.png"
pathScriptFolder = os.path.dirname(pathConfig)

try:    
    # For listening to windows shutdown/sleep/wake up broadcast use.
    winEvent = wmi.WMI().Win32_PowerManagementEvent.watch_for()
    
    # Create a system tray icon.
    sysTray = BuildSystemTray(imgPath)
   
    # Create thread to listen for windows power event to quit script.
    threading.Thread(target=HandleWindowsLatestEvent(winEvent(), sysTray))
   
    # Notify computer squaring has started for what category.
    notif = ToastNotifier()
    notif.show_toast("Auto Powerapps Configurator", f"Setting Up.", duration=5)
    
    # Get today's day and current time.
    datetimeToday = datetime.datetime.today()
    dateToday = datetimeToday.date()
    dayToday = pd.to_datetime(datetimeToday).day_name()
    dayToday = dayToday.lower()
    timeNow = datetimeToday.time()
    print(f"Today: {dayToday}")
    print(f"Time now: {timeNow}")
    print(f"Date: {dateToday}")
    
    # Ensure config file exists.
    if not os.path.exists(pathScriptFolder):
        os.makedirs(pathScriptFolder)
    if not os.path.exists(pathConfig):
        fileConfig = open(pathConfig, "w")
        fileConfig.write("- Email: \n- Password: \n\nCourses:\n")
        for i in range(6):
            fileConfig.write("> \n")
        for i in range(4):
            fileConfig.write("x \n")
            fileConfig.close()
    
    # Ensure template file unmodified.
    fileConfigTemplate = open(pathConfigTemplate, "w")
    fileConfigTemplate.write("- Email: \n- Password: \n\nCourses:\n")
    fileConfigTemplate.write("> Mon, 10:00, TEB2083 Technopreneurship Team Project\n")
    fileConfigTemplate.write("> tues, 12:00, TEB2083/T Technopreneurship Team Project\n")
    fileConfigTemplate.write("> WED, 14:00, TEB2083/Lb Technopreneurship Team Project\n")
    fileConfigTemplate.write("> thuRs, 16:00, GFB2102/GEB2102 (G2) Entrepreneurship\n")
    fileConfigTemplate.write("> frI, 18:00, TEB2103/ Modelling and Simulation\n")
    fileConfigTemplate.write("x sAt, 20:00, TEB2103/T Modelling and Simulation\n")
    fileConfigTemplate.write("x SuN, 22:00, TEB2103/Lb Modelling and Simulation\n")
    for i in range(2):
        fileConfigTemplate.write("> \n")
    for i in range(2):
        fileConfigTemplate.write("x \n")
    fileConfigTemplate.close()
    
    # Ensure run history file exists.
    if not os.path.exists(pathScriptFolder):
        os.makedirs(pathScriptFolder)
    if not os.path.exists(pathRunHistory):
        fileHistory = open(pathRunHistory, "w")
        fileHistory.write("* Auto Powerapps Configurator - Run History *\n")
        fileHistory.close()
    
    # Ensure registered attendances file exists.
    if not os.path.exists(pathScriptFolder):
        os.makedirs(pathScriptFolder)
    if os.path.exists(pathRegisteredAttendnaces):
        fileRegisteredAttendances = open(pathRegisteredAttendnaces, "r")
        if fileRegisteredAttendances.readline() != f"{dateToday}\n":
            fileRegisteredAttendances.close()
            fileRegisteredAttendances = open(pathRegisteredAttendnaces, "w")
            fileRegisteredAttendances.write(f"{dateToday}\n")
            fileRegisteredAttendances.close()
    else:
        fileRegisteredAttendances = open(pathRegisteredAttendnaces, "w")
        fileRegisteredAttendances.write(f"{dateToday}\n")
        fileRegisteredAttendances.close()
    
    # Ensure that Email and Password are provided, and config format is correct.
    fileConfig = open(pathConfig, "r")
    fileLineList = fileConfig.readlines()
    fileConfig.close()
    if fileLineList[0][:9] != "- Email: ":
        raise Exception(f"Error. Wrong email format.\nFirst line first 9 char: '{fileLineList[0][:9]}'")
    if fileLineList[1][:12] != "- Password: ":
        raise Exception(f"Error. Wrong password format.\nSecond line first 12 char: '{fileLineList[1][:12]}'")
    if fileLineList[2] != "\n":
        raise Exception(f"Error. Wrong config format.\nThird line: '{fileLineList[2]}'")
    if fileLineList[3] != "Courses:\n":
        raise Exception(f"Error. Wrong courses format.\nFourth line first 8 char: '{fileLineList[3]}'")
    
    # Retrieve utp email credentials and registering course attendances. 
    userEmail = fileLineList[0].replace("- Email: ", "")
    userPassword = fileLineList[1].replace("- Password: ", "")
    countCourses = len(fileLineList) - 4
    courseList = []
    listSize = len(fileLineList)
    for i in range(4, listSize):
        # Ensure course items format is correct.
        lineCourse = fileLineList[i]
        if lineCourse != "" and lineCourse[0] != ">" and lineCourse[0] != "x":
            raise Exception(f"Error. Wrong course name format.\nLine {i}:'{lineCourse}'")
        commaCount = lineCourse.count(",")
        if commaCount != 0 and commaCount != 2:
            raise Exception(f"Error. Wrong course name comma format.\n:Line {i}:'{lineCourse}'")
        
        # Skip listed course for registration if specified as "x".
        if lineCourse[0] == "x":
            continue
        
        # Retrieve course details from config file.
        lineCourse = lineCourse.replace("> ", "")
        course = lineCourse.split(",") 
        for i in range(len(course)):
            element = course[i].strip()
            course[i] = element
        if course != [""]:
            courseList.append(course)
    
    # Schedule all retrieved course details to register.
    primaryScheduler = schedule.Scheduler()
    for course in courseList:
        day = course[0].upper()
        timing = course[1]
        courseName = course[2]
        timezone = pytz.timezone("Asia/Kuala_Lumpur")
        if day == "MON":
            primaryScheduler.every().monday.at(timing, timezone).do(RegisterAttendance, email = userEmail, password = userPassword, course = courseName)
        elif day == "TUES":
            primaryScheduler.every().tuesday.at(timing, timezone).do(RegisterAttendance, email = userEmail, password = userPassword, course = courseName)
        elif day == "WED":
            primaryScheduler.every().wednesday.at(timing, timezone).do(RegisterAttendance, email = userEmail, password = userPassword, course = courseName)
        elif day == "THURS":
            primaryScheduler.every().thursday.at(timing, timezone).do(RegisterAttendance, email = userEmail, password = userPassword, course = courseName)
        elif day == "FRI":
            primaryScheduler.every().friday.at(timing, timezone).do(RegisterAttendance, email = userEmail, password = userPassword, course = courseName)
        elif day == "SAT":
            primaryScheduler.every().saturday.at(timing, timezone).do(RegisterAttendance, email = userEmail, password = userPassword, course = courseName)
        elif day == "SUN":
            primaryScheduler.every().sunday.at(timing, timezone).do(RegisterAttendance, email = userEmail, password = userPassword, course = courseName)
        
    # List out all the jobs registered.
    print(f"\nRetrieved jobs:\n{primaryScheduler.jobs}\n")
    jobsActiveList = primaryScheduler.jobs
    
    # Retrieve jobs that's supposed to run earlier today and display all registering attendances today.
    passedJobList = []
    passedAttendanceList = []
    jobsActiveTodayList = []
    for job in jobsActiveList:
        if job.start_day == dayToday:
            jobsActiveTodayList.append(job)
            jobRunTime = job.next_run.time()
            if jobRunTime < timeNow:
                passedJobList.append(job)
    print(f"Active jobs today:\n{jobsActiveTodayList}\n")
    print(f"Passed jobs today:\n{passedJobList}\n")
                
    # Retrieve passed job attendances.
    for job in passedJobList:
        passedJobKeywords = [*job.job_func.keywords.values()]  # Unpack dict variable into list of values.
        passedJobCourseName = passedJobKeywords[2]
        passedAttendanceList.append(passedJobCourseName)
    print(f"Passed attendances today:\n{passedAttendanceList}\n")
    
    # Check if all courses has already registered, meaning no jobs left, skip online registered attendances retrieval.
    registeredAttendanceList = []
    fileRegisteredAttendances = open(pathRegisteredAttendnaces, "r")
    fileLineList = fileRegisteredAttendances.readlines()
    listSize = len(fileLineList)
    for i in range(1, listSize):
        lineCourse = fileLineList[i][:-1]
        registeredAttendanceList.append(lineCourse)
    fileRegisteredAttendances.close()
    print(f"Registered attendances today:\n{registeredAttendanceList}\n")
    
    try:
        # Get all registered attendances online, skip if failed. 
        registeredAttendanceList = RetrieveRegisteredAttendances(userEmail, userPassword)
        print(f"Registered attendances today:\n{registeredAttendanceList}\n")
    except Exception as exception:
        # Catch exception.
        errorMessage = str(exception)
        notif.show_toast("Auto Powerapps Configurator", f"Connection error.\nOnline Registered Attendances Retrieval Failed. Script Continue Running.\n" + errorMessage, duration=5)
        print("\nConnection error.\nOnline Registered Attendances Retrieval Failed. Script Continue Running.\n" + errorMessage) 
        
    # Retrieve pending attendance registration jobs.
    pendingJobsTodayList = []
    for job in jobsActiveTodayList:
        pendingJobKeywords = [*job.job_func.keywords.values()]  # Unpack dict variable into list of values.
        pendingJobCourseName = pendingJobKeywords[2]
        if pendingJobCourseName not in registeredAttendanceList:
            pendingJobsTodayList.append(job)
    print(f"Pending attendances today:\n{pendingJobsTodayList}\n")
    
    # Identify and add all not-yet-registered-attendance jobs.
    backupScheduler = schedule.Scheduler()
    notCompletedAttendanceList = [passedAttendance for passedAttendance in passedAttendanceList if passedAttendance not in registeredAttendanceList]
    print(f"Missed attendances today:\n{notCompletedAttendanceList}")
    for job in passedJobList:
        passedJobKeywords = [*job.job_func.keywords.values()]  # Unpack dict variable into list of values.
        passedJobCourseName = passedJobKeywords[2]
        if passedJobCourseName in notCompletedAttendanceList:
            backupScheduler.jobs.append(job)
    backupScheduler.run_all()
     
    if pendingJobsTodayList != []:
        # Notify user .
        notif = ToastNotifier()
        notif.show_toast("Auto Powerapps Configurator", f"Scheduler running.", duration=5)
        print("\nScheduler continuously running for attendance registrations.") 
        
        # Run the scheduler to register courses online overtime.
        while True:
            time.sleep(900)  # 900 seconds = 15 minutes.
            primaryScheduler.run_pending()
    else:
        # Notify user no pending jobs are left today.
        notif = ToastNotifier()
        notif.show_toast("Auto Powerapps Configurator", f"All jobs completed.", duration=5)
        print("\nNo more remaining attendances to register today.") 
            
        # Run infinitely.
        while True:
            time.sleep(86400)  # 86400 seconds = 1 day.
    
except Exception as exception:
    # Catch exception.
    errorMessage = str(exception)
    notif = ToastNotifier()
    notif.show_toast("Auto Powerapps Configurator", f"Operation Failed.\n" + errorMessage, duration=5)
    print("\nAuto Powerapps Configurator script terminated. Operation failed.\n" + errorMessage) 
    
    
# End script properly if it reached the end of script for any reason.
try:
    QuitScript(sysTray)
    print("Script ended. Script terminating.")
except NameError:
    print("Systray has not been build. Script ended. Script terminating.")