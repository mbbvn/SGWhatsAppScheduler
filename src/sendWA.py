from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pyperclip
import time
import sys
from datetime import datetime 
from pathlib import Path
import logging

logger = logging.getLogger(__name__)

groupDict = {
    "TEST_GROUPS": "test_groups.txt",
    "ALL_GROUPS" : "all_groups.txt",
    "JAPA_GROUPS" : "japa_groups.txt",
    "ENG_GROUPS" : "eng_groups.txt",
    "ENG_REPEAT_GROUPS" : "eng_groups_repeat_if_failed.txt",
    "HINDI_GROUPS" : "hindi_groups.txt",
    "GMC_B1_THU_GROUPS" : "gmc_b1_thu_groups.txt",
    "GMC_B1_SAT_GROUPS" : "gmc_b1_sat_groups.txt",
    "GMC_B1_SUN_GROUPS" : "gmc_b1_sun_groups.txt",
    "GMC_B2_SAT_GROUPS" : "gmc_b2_sat_groups.txt",
    "GMC_B1_ALL_GROUPS" : "gmc_b1_all_groups.txt",
    "GMC_ALL_GROUPS" : "gmc_all_groups.txt",
    "CC_ALL_GROUPS" : "cc_all_groups.txt",
    "GME_ALL_GROUPS" : "gme_all_groups.txt"
}

# sendMsgResult = {
# 	"msgStatus"    : "", #SUCCESSFUL, FAILED
# 	"sentToGrps"   : "", #comma-separated group names
# 	"failedToGrps" : ""  #commad-separate group names
# }

#configPath = Path(Path().absolute(), "..\config")
projectRoot = Path(__file__).parents[1]
imgPath = Path(projectRoot, "img")
configPath = Path(projectRoot,"config")
picOnlyMsg = False

def invoke_wahandler(msgDict):
    logger.info("invoke_wahandler")
    global groupDict
    logger.debug(groupDict)
    
    if(is_input_data_good(msgDict)):
        #proceed to send the WA message
        msgName = msgDict.get("msgName")
        msgText = msgDict.get("msgText")
        msgGroupAlias = msgDict.get("msgGroupAlias").upper()
        groupsfile = groupDict[msgGroupAlias]
        logger.info("groupsfile: %s",groupsfile)
        picPath = msgDict.get("msgPic")
        #check whether the message has a pic
        hasPic = True if picPath is not None else False
        if(hasPic):
            logger.debug("picPath: %s", picPath)
            return sendWhatsApp(msgText,groupsfile,picPath)
        else:
            return sendWhatsApp(msgText,groupsfile)

    else:
        #data issue. Cannot send WA message
        logger.error("data issue")
    
def is_input_data_good(msgDict):
    logger.info("checking input data")
    global picOnlyMsg
    #print (msgDict)
    #msgText and msgGroupAlias cannot be empty/none
    if (msgDict.get("msgGroupAlias") is None): 
        logger.error("Input data is not okay msgGroupAlias is missing")
        return False
    if(msgDict.get("msgText") is None): 
        if (msgDict.get("msgPic") is not None): 
            logger.info("Input data only has pic and no message")
            picOnlyMsg = True
            return True
        else:
            logger.error("Input data is not okay. msgText and msgPic are missing")
            return False
    else:
        logger.info("Input data is okay")
        return True

def sendWhatsApp(msg, groupsFile, pic=None):
    sendMsgResult = dict.fromkeys(['msgStatus', 'sentToGrps', 'failedToGrps'])
    sentToGrps = ""
    failedToGrps = ""
    try:
        if groupsFile:
            with open(Path(configPath,groupsFile), 'r', encoding='utf8') as f:
                groups = [group.strip() for group in f.readlines()]
                logger.info("Groups this mesage will be sent to: %s", groups)
    except IndexError:
        logger.error('Please provide the group name as the first argument.')
        sendMsgResult['msgStatus'] = "FAILED"
        return sendMsgResult
    if(pic is not None): #log picture details
        logger.debug("pic: %s", pic)
        absPicPath=str((Path(imgPath,pic)))
        logger.info("Picture path: %s", absPicPath)

    options = webdriver.ChromeOptions()
    #options.add_argument(CHROME_PROFILE_PATH)
    #Chrom version 103 introduced an issue on 1-Jul-2022. Fix is expected in Chrome v104. 
    #   Until then the beta version is being used. Below line can be removed after the v104 is installed
    #options.binary_location = "C:/Program Files/Google/Chrome Beta/Application/chrome.exe"
    options.add_argument("user-data-dir=C:\\Users\\automation\\appdata\\Local\\Google\\Chrome\\UserData\\Default")

    browser = webdriver.Chrome(
        #executable_path='/chromedriver_96.0.4664.45_win32/chromedriver.exe', options=options)
        executable_path='/chromedriver_win32/chromedriver.exe', options=options)

    browser.maximize_window()

    browser.get('https://web.whatsapp.com/')
    #this is for real..not for testing..wait 30 seconds for the page to load fully 
    #19-Nov-2023 Observed that Chat sync in web whatsapp is taking longer, particularly for the morning japa msg. This further causes send failures as some of the groups cannot be  found in the search bar until the chat sync is complete. Increasing the wait time to 120 sec/2 mins seem to have helped. No changes made at this time though. Observe for few days and change. This issue has been going on for more than a month now.
    #20-Nov-2023 Increased the sleep time by another 60 seconds to allow Chat Sync. However, it appears to be needed only for the morning Japa msg. But causes additional 1 min delay for all schedules (ARK and evening japa msg). Need to be revised as needed.
    #26-Dec-2023 Changing the sleep time to 90 seconds
    #28-Dec-2023 Chat sync is still taking more time and causing msg failures, particularly for the morning Japa msg. Hence changing it to 180 seconds (3 min)
    time.sleep(180) 
    #time.sleep(300) #uncomment only for testing

    for group in groups:
        msgDeliveryFlag = False # reset to false for each iteration
        logger.info("Attempting to send message to group: %s", group)
        try:
            search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'

            #search_box = WebDriverWait(browser, 300).until(
            search_box = WebDriverWait(browser, 300).until(
                EC.presence_of_element_located((By.XPATH, search_xpath))
            )

            search_box.clear()

            time.sleep(1)

            pyperclip.copy(group)

            search_box.send_keys(Keys.CONTROL + "v")

            time.sleep(2)

            group_xpath = f'//span[@title="{group}"]'
            group_title = browser.find_element_by_xpath(group_xpath)

            group_title.click()

            time.sleep(1)

        except: # catch-all
            logger.exception("Fatal error while finding group", exc_info=True)
            #populate sendMsgResult
            sendMsgResult['msgStatus'] = "FAILED"
            failedToGrps += str(group) + ","
            continue # continue sending message to other groups

        try:
            if (pic is not None): # if picture is present
                # Find the paper clip (attach) element 
                attachment_box = browser.find_element_by_xpath(
                    '//div[@title="Attach"]')
                attachment_box.click()
                time.sleep(3) # allow the animation to load
                
                # find 'attach image' element
                image_box = browser.find_element_by_xpath(
                    '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
                image_box.send_keys(absPicPath)
                time.sleep(3)
                
                # find text element to paste the text msg
                #input_xpath = '//div[@contenteditable="true"][@data-tab="9"]'
                input_xpath = '//div[@contenteditable="true"][@data-tab="10"]'
                input_box = browser.find_element_by_xpath(input_xpath)
                # copy the message text to clip board
                if(picOnlyMsg is False):
                    pyperclip.copy(msg)
                    logger.debug("~~~~~~~~~~~~ Message to send....BEGIN ~~~~~~~~~~~~~~~")
                    logger.debug(msg)
                    logger.debug("~~~~~~~~~~~~ Message to send....END ~~~~~~~~~~~~~~~")
                    #Ctrl+V to paste the message into text box
                    input_box.send_keys(Keys.CONTROL + "v")
                    time.sleep(1)
                #locate the Send button
                send_btn = browser.find_element_by_xpath(
                    '//span[@data-icon="send"]')
                logger.info("Sending with picture")
                #time.sleep(1) # allow links to load thumbnails, if any
                time.sleep(5) # allow links to load thumbnails, if any
                #Click Send
                send_btn.click()
                time.sleep(5) #increasing the wait time from 2 to 5 seconds as some times image is not posted immediately
                msgDeliveryFlag = True
            else: # no picture
                time.sleep(2)
                #input_xpath = '//div[@contenteditable="true"][@data-tab="9"]'
                input_xpath = '//div[@contenteditable="true"][@data-tab="10"]'
                #time.sleep(300)#TODO remove this line

                input_box = browser.find_element_by_xpath(input_xpath)

                pyperclip.copy(msg)
                input_box.send_keys(Keys.CONTROL + "v") 
                logger.info("Sending text-only message")
                time.sleep(3) # allow links to load thumbnails, if any
                input_box.send_keys(Keys.ENTER)
                time.sleep(3)
                msgDeliveryFlag = True
        except IndexError:
            pass #FIXME why pass?
        
        except: # catch-all
            logger.exception("Fatal error while sending WhatsApp Message", exc_info=True)
            continue # continue sending message to other groups

        finally:
            #FIXME any resources to cleanup?
            if(msgDeliveryFlag == False):
                logger.error("Failed to send message to group %s", group)
                failedToGrps += str(group) + ","

            else:
                logger.info("Succesfully sent message to group %s", group)
                sentToGrps += str(group) + ","
    
    # Populate sendMsgResult dictionary      
    if(len(failedToGrps) > 1):
        logger.error("Message has not been delivered to some or all groups. Check failedToGrps value for more info.")
        sendMsgResult['msgStatus'] = "FAILED"
    
    elif(len(sentToGrps) > 1):
        logger.info("Message has been delivered to all groups")
        sendMsgResult['msgStatus'] = "SUCCESSFUL"
    
    else:
        logger.error("Message has not been delivered to any groups")
        sendMsgResult['msgStatus'] = "FAILED"

    sendMsgResult['failedToGrps'] = failedToGrps
    sendMsgResult['sentToGrps'] = sentToGrps

    return sendMsgResult