import pandas as pd
from config.config import *
import logging
import xlwings as xw

logger = logging.getLogger(__name__)


def load_detailedSheet():
    workbook = xw.Book(DtlShtPath)
    return workbook

def read_detailedSheet():
    # Read and filter the detailed sheet
    DtlShtData = pd.read_excel(DtlShtPath, DtlShtName)
    logger.info(f"Loaded {DtlShtData.shape} rows from the detailed sheet.")
    return DtlShtData

def get_sorted_DtlShtData(DtlShtData):
    # Keep specific columns and filter rows with 'Awaited' status
    columns_to_keep = ['RFP Name', 'Status', 'Per No', 'Hours', 'Start Date', 'End Date', 'NN', 'SupCmpCode',
                       'CRM ID', 'RowNum', 'BidTrigNN', 'FioriNuM', 'FioriCreDate','Act', 'FioriApprover']
    DtlShtData = DtlShtData[columns_to_keep]
    DtlShtData = DtlShtData[DtlShtData['Status'].str.contains('Awaited')]
    logger.info(f"Filtered {DtlShtData.shape} rows with 'Awaited' status.")

    # Creating two columns "BidTrigNN" and "PerRow
    DtlShtData['BidTrigNN'] = DtlShtData['RFP Name'].str[-5:] + '-' + DtlShtData['SupCmpCode'] + '-' + DtlShtData[
        'NN'].astype(str)
    DtlShtData['PerRow'] = DtlShtData['Per No'].astype(str) + '-' + DtlShtData['RowNum'].astype(str)

    # extracting rows where "FioriNuM" is empty means fiori generation is pending
    empty_rows = DtlShtData[DtlShtData['FioriNuM'].isnull()]
    sorted_DtlShtData = empty_rows.sort_values(by='BidTrigNN').reset_index(drop=True)
    logger.info(f"Number of rows needed fiori generation: {len(sorted_DtlShtData)}")

    # Creating fiori dates from start date and end date
    sorted_DtlShtData['Start Date'] = pd.to_datetime(sorted_DtlShtData['Start Date'], format='%m/%d/%Y').dt.strftime(
        '%d.%m.%Y')
    sorted_DtlShtData['End Date'] = pd.to_datetime(sorted_DtlShtData['End Date'], format='%m/%d/%Y').dt.strftime(
        '%d.%m.%Y')
    sorted_DtlShtData['FioriDate'] = sorted_DtlShtData['Start Date'] + " to " + sorted_DtlShtData['End Date']

    # Creating rule for customzied Proposal Name
    def PropName(Prop):
        if len(Prop) > 35:
            return Prop[:31] + '_' + Prop[-5:]
        else:
            return Prop

    sorted_DtlShtData['Description'] = sorted_DtlShtData['RFP Name'].apply(PropName)
    return sorted_DtlShtData




def get_BidTrigNN_list(sorted_DtlShtData):
    BidTrigNN_list = sorted_DtlShtData['BidTrigNN'].drop_duplicates().tolist()
    logger.info(f"Number of unique BidTrigNN or number of fiori should be created: {len(BidTrigNN_list)}")
    return BidTrigNN_list

def get_df_RowNumBidTrigNN(sorted_DtlShtData):
    df_RowNumBidTrigNN = sorted_DtlShtData[['RowNum', 'BidTrigNN']]
    return df_RowNumBidTrigNN

def get_FioriNumbermap(BidTrigNN_list,FioriNumber,df_RowNumBidTrigNN):
    FioriNumbermap = pd.DataFrame({'BidTrigNN': BidTrigNN_list, 'FioriNuM': FioriNumber})
    FioriNumbermap = pd.merge(df_RowNumBidTrigNN, FioriNumbermap, on=['BidTrigNN'], how='left')
    FioriNumbermap=FioriNumbermap.sort_values(by='RowNum').reset_index(drop=True)
    return FioriNumbermap

#================== Merging DtlShtData and FioriNumbermap ==================================================
def get_DtlShtDataFiori(DtlShtData,FioriNumbermap):
    DtlShtDataFiori = pd.merge(DtlShtData, FioriNumbermap, on=['RowNum'], how='left',
                               suffixes=('_df1', '_df2'))
    DtlShtDataFiori['FioriNuM'] = DtlShtDataFiori['FioriNuM_df1'].replace("", pd.NA)  # Handle empty strings as NaN
    DtlShtDataFiori['FioriNuM'] = DtlShtDataFiori['FioriNuM_df1'].combine_first(DtlShtDataFiori['FioriNuM_df2'])
    DtlShtDataFiori = DtlShtDataFiori.drop(columns=['FioriNuM_df1','FioriNuM_df2'])  # Drop the extra column from the merge
    DtlShtDataFiori = DtlShtDataFiori.sort_values(by='RowNum').reset_index(drop=True)
    return DtlShtDataFiori


#================== Pasting fiori number/status in detailed sheet==================================

def write_to_excel(Column,sheet_path, sheet_name, start_row=2, start_col=2):
    wb = xw.Book(sheet_path)
    sheet = wb.sheets[sheet_name]
    sheet.range((start_row, start_col)).options(transpose=True).value = Column.tolist()
    wb.save()
    wb.close()
    logger.info(f"Data written successfully to {sheet_name}")


