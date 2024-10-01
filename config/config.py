# File paths and URL
DtlShtPath = r"C:\Users\enesine\Downloads\MOAI Tracker\Fiori_Automation\Update_Detailsheet_MELA.xlsm"
DtlShtName = "Detailed_Sheet"
URL = "https://erpbusinessapps.ericsson.net/sap/bc/ui2/flp#ZFIRRB-display&/NewSAFRequest"
UserName = "hello@gmail.com"
Password = "hello123"
webdriver_path = r"C:\Users\enesine\PycharmProjects\selenium_fiori_automation\drivers\msedgedriver.exe"
backup_path = r"C:\Users\enesine\Downloads\MOAI Tracker\Fiori_Automation"
FioriNumber = []
log_file_path=r"C:\Users\enesine\PycharmProjects\selenium_fiori_automation\logs\application.log"

# Selenium locators
Inputbox_Description_xpath = "//*[contains(@id,'idDescription')]/input"
Inputbox_NWID_xpath = "//*[contains(@id,'idCopyCostObjectOrdering')]/input"
Element_Loading_xpath1="//*[@id='application-ZFIRRB-display-component---NewSAFRequest-busyIndicator']/div"
Element_Loading_xpath="//*[@id='__page0-busyIndicator']/div"
Inputbox_ActivityID_xpath = "//*[contains(@id,'idNavNetwToActvty')]/input"
CheckBox_idNoNewActivity_xpath = "//*[contains(@id,'idNoNewActivity-CbBg')]"
Inputbox_SLC_xpath = "//*[contains(@id,'idSupplyingCompanyInput')]/input"
Button_ContButton_xpath = "//*[@id='__button10-BDI-content']"
Button_NextPer_xpath = "//*[@id='__button1-inner']"
DropDown_ApproverDropDown_xpath = "//*[contains(@id,'idApprover-label') and contains(@class,'sapMSltLabel')]"
Link_ApproverNames_xpath = "//span[@style='width: 40%;']"
Textbox_Comment_xpath = "//*[contains(@id,'idComment')]/textarea"
Button_SaveAsDraft_xpath = "//*[@id='__button16-BDI-content']"
Button_Submit_xpath = "//*[@id='__button17-BDI-content']"
Text_FioriNo_xpath = "//span[contains(text(), 'Form successfully')]"
Button_CreateNewSAF_xpath = "//bdi[contains(text(), 'Create New SAF')]"
Text_approverPage_xpath = "//*[contains(text(),'agreement')]"


