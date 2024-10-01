from selenium.webdriver import Keys
from selenium.webdriver.common.by import By

from services.data_processing import *
from services.selenium_service import *
import logging
import time
from utils.logging_setup import setup_logger

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize the logger
# logger = setup_logger()

# Prevents sleep
prevent_sleep()
def main():
    logger.info("Script started.")

    # Load and filter data
    dtlShtData = read_detailedSheet()

    # get_sorted_DtlShtData
    sorted_DtlShtData = get_sorted_DtlShtData(dtlShtData)

    # get_df_RowNumBidTrigNN
    df_RowNumBidTrigNN=get_df_RowNumBidTrigNN(sorted_DtlShtData)

    # get_BidTrigNN_list
    BidTrigNN_list=get_BidTrigNN_list(sorted_DtlShtData)

    # Initialize WebDriver
    driver = start_selenium()

    # Iterating BidTrigNN_list
    for item in BidTrigNN_list:
        logger.info(f"Processing BidTrigNN id: {item}")
        filtered_sorted_DtlShtData = sorted_DtlShtData[sorted_DtlShtData['BidTrigNN'] == item]
        NoOfPer = int(filtered_sorted_DtlShtData['Per No'].size)
        logger.info(f"Number of Person in this BidTrigNN id: {NoOfPer}")

        Des = filtered_sorted_DtlShtData['Description'].iloc[0]
        NN = filtered_sorted_DtlShtData['NN'].astype(str).iloc[0]
        SLC = filtered_sorted_DtlShtData['SupCmpCode'].iloc[0]
        Approver_name = filtered_sorted_DtlShtData['FioriApprover'].iloc[0]
        CRMID = filtered_sorted_DtlShtData['CRM ID'].iloc[0]
        if SLC == '255':
            SLC = '2828'
        elif SLC == '116':
            SLC = 'TEM'
        try:
            driver.find_element(By.XPATH, Inputbox_Description_xpath).send_keys(Des)
            driver.find_element(By.XPATH, Inputbox_NWID_xpath).send_keys(NN, Keys.ENTER)
            # time.sleep(10)
            WebDriverWait(driver, 20).until(EC.invisibility_of_element_located((By.XPATH, Element_Loading_xpath)))
            time.sleep(2)
            driver.find_element(By.XPATH, CheckBox_idNoNewActivity_xpath).click()
            driver.find_element(By.XPATH, Inputbox_SLC_xpath).send_keys(SLC, Keys.ENTER)
            time.sleep(5)

            # =====================iterating for each Per no in an ICRRB===============================================

            for i in range(NoOfPer):
                Inputbox_PerDropDown_xpath = f'//*[@id="__input2-application-ZFIRRB-display-component---NewSAFRequest--idManHours-{i}-inner"]'
                DropDown_RemoteSelect_xpath = f'//*[@id="__select0-application-ZFIRRB-display-component---NewSAFRequest--idManHours-{i}-label"]'
                Link_RemoteSelect1_xpath = f'//li[contains(@id,"idManHours-{i}") and contains(text(), "Remote")]'
                Date_DateSelect_xpath = f'//*[@id="__selection0-application-ZFIRRB-display-component---NewSAFRequest--idManHours-{i}-inner"]'
                Input_HoursSelect_xpath = f'//*[@id="__input3-application-ZFIRRB-display-component---NewSAFRequest--idManHours-{i}-inner"]'

                PerNo = filtered_sorted_DtlShtData['Per No'].astype(str).tolist()
                Hrs = filtered_sorted_DtlShtData['Hours'].tolist()
                fdate = filtered_sorted_DtlShtData['FioriDate'].tolist()

                logger.info(
                    f"Processing: Description={Des}, NN={NN}, SLC={SLC}, PerNo={PerNo[i]}, Hrs={Hrs[i]}, Date={fdate[i]}")

                driver.find_element(By.XPATH, Inputbox_PerDropDown_xpath).send_keys(PerNo[i], Keys.ENTER)
                time.sleep(2)
                try:
                    driver.find_element(By.XPATH, DropDown_RemoteSelect_xpath).click()
                    driver.find_element(By.XPATH, Link_RemoteSelect1_xpath).click()
                except Exception as exp:
                    pass
                driver.find_element(By.XPATH, Date_DateSelect_xpath).send_keys(fdate[i])
                time.sleep(2)
                driver.find_element(By.XPATH, Input_HoursSelect_xpath).send_keys(str(round(Hrs[i])))
                time.sleep(2)
                driver.find_element(By.XPATH, Button_NextPer_xpath).click()
                time.sleep(5)
            driver.find_element(By.XPATH, Button_ContButton_xpath).click()
            time.sleep(2)
            driver.find_element(By.XPATH, Text_approverPage_xpath).click()
            logger.info(f"Driver moved to approver page")
            driver.find_element(By.XPATH, DropDown_ApproverDropDown_xpath).click()
            time.sleep(2)
            Approver_names = driver.find_elements(By.XPATH, Link_ApproverNames_xpath)

            for element in Approver_names:
                if element.text == Approver_name:
                    logger.info(f"Approver name found")
                    element.click()
                    break
            else:
                raise ValueError(f"No match found for {Approver_name}")

            time.sleep(2)
            driver.find_element(By.XPATH, Textbox_Comment_xpath).send_keys("Salesforce ID: " + str(CRMID))

            logger.info(f"Pasted Salesforce ID:: {str(CRMID)}")

            driver.find_element(By.XPATH, Button_SaveAsDraft_xpath).click()
            time.sleep(2)
            fiorinum = driver.find_element(By.XPATH, Text_FioriNo_xpath).text

            logger.info(f"FioriNumber created: {fiorinum[-11:]}")

            driver.find_element(By.XPATH, Button_CreateNewSAF_xpath).click()
            time.sleep(5)
            FioriNumber.append("SAF"+fiorinum[-11:])

        except Exception as e:
            if 'idNoNewActivity' in str(e) or 'interactable' in str(e):
                print(f"Network ID is incorrect for {item}")
                FioriNumber.append(f"Network ID is incorrect for {item}")
                driver.get(URL)
                WebDriverWait(driver, 30).until(EC.title_contains("Intercompany"))
            elif 'agreement' in str(e):
                 print(f"Personnel No is incorrect for {item}")
                 FioriNumber.append(f"Personnel No is incorrect for {item}")
                 driver.get(URL)
                 WebDriverWait(driver, 30).until(EC.title_contains("Intercompany"))
            elif '__button1-inner' in str(e):
                print(f"Hourly rate is not maintained for {item}")
                FioriNumber.append(f"Hourly rate is not maintained for {item}")
                driver.get(URL)
                WebDriverWait(driver, 30).until(EC.title_contains("Intercompany"))
            elif 'No match' in str(e) or 'agreement' in str(e):
                print(f"Approver name is incorrect or in small character for {item}: {e}")
                FioriNumber.append(f"Approver name is incorrect or in small character for {item}: {e}")
                driver.get(URL)
                WebDriverWait(driver, 30).until(EC.title_contains("Intercompany"))
            else:
                print(f"Any other error for {item}: {e}")
                FioriNumber.append(f"Any other error for {item}: {e}")
                driver.get(URL)
                WebDriverWait(driver, 30).until(EC.title_contains("Intercompany"))

            driver.refresh()
            time.sleep(2)

        logger.info(f"Completed processing for BidTrigNN: {item}")
        logger.info(f"#################################################################################")
        logger.info(f"#################################################################################")

    fioriNumbermap = get_FioriNumbermap(BidTrigNN_list, FioriNumber, df_RowNumBidTrigNN)
    pd.set_option('display.max_columns', None)
    print(fioriNumbermap)
    DtlShtDataFiori=get_DtlShtDataFiori(dtlShtData, fioriNumbermap)

    Column=DtlShtDataFiori['FioriNuM']
    sheet_path = DtlShtPath
    sheet_name = 'Detailed_Sheet'

    write_to_excel(Column, sheet_path, sheet_name, start_row=2, start_col=35)

    show_popup()

    driver.quit()
    logger.info("Script ended successfully.")


if __name__ == "__main__":
    main()
restore_sleep()  # Restores sleep after tests