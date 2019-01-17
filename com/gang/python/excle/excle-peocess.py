
import pandas as pd

import openpyxl

#无效资产证明.xls	有效资产证明.xls
invalid_fileName = '/Users/ligang/doc/无效资产证明(1).xls'
valid_fileName = '/Users/ligang/doc/有效资产证明(1).xls'

filePath = '/Users/ligang/doc/'

def readExcle(fileName):
    df = pd.read_excel(fileName)
    df.columns = ['index_num', 'TX_ACCT_NO', 'ID', 'HBONE_NO', 'AMOUNT', 'FILE_NO',
       'AUDIT_STATUS', 'AUDIT_DATE', 'TERM_START', 'TERM_END', 'REMARK',
       'CERTIFICATE_TYPE', 'OUTLET_CODE', 'VERSION', 'CREATE_DATE',
       'UPDATE_DATE', 'INVALID', 'INVALID_DATE', 'SOURCE', 'USER_TYPE',
       'IMPORT_DATA', 'GENERATE_FROM']


    dict = {}
    for row in df.itertuples():
        txAcctNo = getattr(row, "TX_ACCT_NO")
        if(txAcctNo in dict):
            dict.get(txAcctNo).append(row)
        else:
            dict[txAcctNo] = [row]
    return dict

Person_user_type = 1
part_pass_status = 2
all_pass_status = 3
income_type = 1
curr_dt = 20190117
form_sys_type = 1
hb_source = 2
out_source = 1
need_amount = 3000000
new_file_invalid = 'invalid_1.xlsx'
new_file_valid = 'valid_1.xlsx'
def processStatus(inputFile, outFile):
    dict = readExcle(inputFile)
    rsList = []
    for info in dict.values():
        rst = processCustStatus(info)
        rsList.append(rst)
    rsdata = pd.DataFrame(rsList)
    print(rsdata)
    newFile = filePath + outFile
    rsdata.to_excel(newFile)


def processCustStatus(info):
    totalAmount = 0
    validateFlag = 0
    TX_ACCT_NO = ''
    for cust in info:
        TX_ACCT_NO = getattr(cust, 'TX_ACCT_NO')
        AUDIT_STATUS = getattr(cust, 'AUDIT_STATUS')
        AMOUNT = getattr(cust, 'AMOUNT')
        TERM_END = getattr(cust, 'TERM_END')
        CERTIFICATE_TYPE = getattr(cust, 'CERTIFICATE_TYPE')
        INVALID = getattr(cust, 'INVALID')
        USER_TYPE = getattr(cust, 'USER_TYPE')
        GENERATE_FROM =getattr(cust, 'GENERATE_FROM')
        SOURCE = getattr(cust, 'SOURCE')
        if audit_pass(AUDIT_STATUS, TERM_END):
            if Person_user_type != USER_TYPE:
                if income_type == CERTIFICATE_TYPE:
                    validateFlag = 1
            else:
                if income_type == CERTIFICATE_TYPE:
                    validateFlag = 1
                else:
                    validateSource = validSouce(SOURCE, GENERATE_FROM)
                    if validateSource:
                        totalAmount = totalAmount + AMOUNT

    status = validateFlag == 1 or (totalAmount >= need_amount)

    status_flag = '0'
    if status:
        status_flag = '1'

    dict = {'txAcctNo': TX_ACCT_NO, 'status':status_flag, 'totalAmt': totalAmount}
   # print(dict)
    return dict



def validSouce(SOURCE, GENERATE_FROM):
    return out_source == SOURCE or (hb_source == SOURCE and form_sys_type == GENERATE_FROM)


def audit_pass(status, termEnd):
    return (all_pass_status == status or part_pass_status == status) and (termEnd >= curr_dt)

processStatus(invalid_fileName, new_file_invalid)

processStatus(valid_fileName, new_file_valid)

