import os
import shutil
from datetime import date
from mysql.connector import pooling
import cx_Oracle
import xlrd
from xlwt import Workbook

xlrd.xlsx.ensure_elementtree_imported(False, None)
xlrd.xlsx.Element_has_iter = True

today = date.today()
today_date = today.strftime("%d/%m/%Y")

user = 'BRUKDB'
passw = 'at1Lord_572'
sid = 'pssprod1'
ip = '10.1.1.51'
port = '1521'


def create_mysql_connection():
    connection_pool = pooling.MySQLConnectionPool(pool_name="pynative_pool",
                                                  pool_size=5,
                                                  pool_reset_session=True,
                                                  host='localhost',
                                                  database='settle',
                                                  port='3308',
                                                  user='root',
                                                  password='root')
    connection_object = connection_pool.get_connection()
    mycursor = connection_object.cursor(buffered=True)
    return mycursor,connection_object,connection_pool


def close_mysql_connection(connection_object,mycursor):
    if connection_object.is_connected():
        mycursor.close()
        connection_object.close()
    else:
        pass


def create_oracle_connection():
    dsn_tns = cx_Oracle.makedsn(ip, port, sid=sid)
    pool = cx_Oracle.SessionPool(user, passw, dsn=dsn_tns, min=2, max=5, increment=1, encoding="UTF-8")
    connection = pool.acquire()
    cursor = connection.cursor()
    return cursor,connection,pool


def close_oracle_connection(pool,connection):
    pool.release(connection)
    pool.close()


def get_pan(c,rrn,amountt,hst):
    rrn = str(rrn)
    amountt = str(amountt)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7) from authorization where aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                                             "select substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7) from authorization_hst where aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7) from authorization where aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            pass
        return row


def get_amount(c,rrn,pann,hst):
    rrn = str(rrn)
    pann = str(pann)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select aut_tran_amou_f004 from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                                              "select aut_tran_amou_f004 from authorization_hst where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select aut_tran_amou_f004 from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            pass
        return row


def get_amount_and_pan(c,rrn,hst):
    rrn = str(rrn)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7) from authorization where aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                   "select aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7) from authorization_hst where aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7) from authorization where aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            pass
        return row


def get_bank(bank_code):
    bank_code_result = str(bank_code)
    if bank_code_result == '8671':
        bankk= 'UB'
    elif bank_code_result == '7601':
        bankk= 'NIB'
    elif bank_code_result == '841':
        bankk= 'AIB'
    elif bank_code_result == '7641':
        bankk= 'ADIB'
    elif bank_code_result == '7661':
        bankk= 'CBO'
    elif bank_code_result == '7621':
        bankk= 'BRIB'
    else:
        bankk = 'BANK_NOT_FOUND'
    return bankk


def check_authorization(c,rrn,pann,amountt,hst):
    rrn = str(rrn).strip()
    pann = str(pann).strip()
    amountt = str(amountt).strip()
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012,aut_acq_bank_code from authorization where aut_sour_inte_code = '16' and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                                                                                                                                                                                                                                                        "select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012,aut_acq_bank_code from authorization_hst where aut_sour_inte_code = '16' and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012,aut_acq_bank_code from authorization where aut_sour_inte_code = '16' and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row

def check_authorization_for_atm(c,rrn,pann,amountt,hst):
    rrn = str(rrn).strip()
    pann = str(pann).strip()
    amountt = str(amountt).strip()
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012,aut_acq_bank_code from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                                                                                                                                                                                                                          "select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012,aut_acq_bank_code from authorization_hst where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012,aut_acq_bank_code from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def check_pos_transaction(c,rrn,pann,amountt,hst):
    rrn = str(rrn)
    pann = str(pann)
    amountt = str(amountt)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012 from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                                                                                                                                                                                                        "select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012 from authorization_hst where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012 from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def check_amount(c,rrn,pann,hst):
    rrn = str(rrn)
    pann = str(pann)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012 from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                                                                                                                                                           "select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012 from authorization_hst where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select aut_retr_ref_numb_f037,aut_resp_code_f039,aut_tran_amou_f004,substr(aut_prim_acct_numb_f002,1,6)||'******'|| substr(aut_prim_acct_numb_f002,13,7),aut_sour_inte_code,aut_date_time_tran_f012 from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def check_reversal_status(c,rrn,pann,amountt,hst):
    rrn = str(rrn)
    pann = str(pann)
    amountt = str(amountt)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select aut_date_time_tran_f012 from authorization where (aut_reve_stat = 'Y' OR aut_reve_stat = 'F') and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                                                                                                                                                 "select aut_date_time_tran_f012 from authorization_hst where (aut_reve_stat = 'Y' OR aut_reve_stat = 'F') and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select aut_date_time_tran_f012 from authorization where (aut_reve_stat = 'Y' OR aut_reve_stat = 'F') and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def check_pre_aut_status(c,rrn,pann,amountt,hst):
    rrn = str(rrn)
    pann = str(pann)
    amountt = str(amountt)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select aut_date_time_tran_f012 from authorization where aut_mess_reas_code_f025 = '1655' and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                                                                                                                                                 "select aut_date_time_tran_f012 from authorization_hst where (aut_mess_reas_code_f025 = '1655') and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select aut_date_time_tran_f012 from authorization where aut_mess_reas_code_f025 = '1655' and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row

def check_ipm_transaction_type(c,rrn,pann,amountt,hst):
    # CIS/MDS
    rrn = str(rrn)
    pann = str(pann)
    amountt = str(amountt)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select aut_date_time_tran_f012 from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_dest_inte_code in (\'61\',\'74\',\'60\',\'62\',\'66\') and aut_retr_ref_numb_f037 = \'" + rrn + "\' union "
                                                                                                                                                                                                                                                                 "select aut_date_time_tran_f012 from authorization_hst where (aut_reve_stat = 'Y' OR aut_reve_stat = 'F') and aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    else:
        c.execute("select aut_date_time_tran_f012 from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_dest_inte_code in (\'61\',\'74\',\'60\',\'62\',\'66\') and aut_retr_ref_numb_f037 = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def check_tran_transaction_record(c,institution,rrn, pann, amountt,hst):
    rrn = str(rrn)
    pann = str(pann)
    amountt = str(amountt)
    hst = str(hst)
    if institution == 'ET_Switch':
        c.execute("select itr_proc_date,itr_acqu_refe_numb from svb_transaction_record where itr_tran_amou = \'" + amountt + "\' and itr_acco_numb like \'%" + pann + "\' and (itr_retr_refe_numb = \'" + rrn + "\' or itr_icc_issu_auth_data = \'" + rrn + "\')")
    elif institution == 'DCI':
        c.execute("select dtr_proc_date,dtr_reference_number from dci_transaction_record WHERE dtr_approval_code in (select aut_auth_id_resp_f038 from authorization where aut_prim_acct_numb_f002 like \'%" + pann + "\' and aut_tran_amou_f004 = \'" + amountt + "\' and aut_retr_ref_numb_f037 = \'" + rrn + "\')")

    elif institution == 'MasterCard':
        if hst == 'Yesforhst':
            c.execute("select itr_proc_date,itr_acqu_refe_numb from ipm_transaction_record where itr_tran_amou = \'" + amountt + "\' and itr_acco_numb like \'%" + pann + "\' and (itr_icc_issu_auth_data = \'" + rrn + "\' or itr_retr_refe_numb = \'" + rrn + "\') union "
                                                                                                                                                                                                                                                                "select itr_proc_date,itr_acqu_refe_numb from ipm_transaction_record_hst where itr_tran_amou = \'" + amountt + "\' and itr_acco_numb like \'%" + pann + "\' and (itr_icc_issu_auth_data = \'" + rrn + "\' or itr_retr_refe_numb = \'" + rrn + "\')")
        else:
            c.execute("select itr_proc_date,itr_acqu_refe_numb from ipm_transaction_record where itr_tran_amou = \'" + amountt + "\' and itr_acco_numb like \'%" + pann + "\' and (itr_icc_issu_auth_data = \'" + rrn + "\' or itr_retr_refe_numb = \'" + rrn + "\')")
    elif institution == 'VISA':
        if hst == 'Yesforhst':
            c.execute("select vtr_proc_date,vtr_acqu_refe_numb from visa_transaction_record where vtr_sour_amou = \'" + amountt + "\' and vtr_acco_numb like \'%" + pann + "\' and vtr_res4 = \'" + rrn + "\' union "
                                                                                                                                                                                                          "select vtr_proc_date,vtr_acqu_refe_numb from visa_transaction_record_hst where vtr_sour_amou = \'" + amountt + "\' and vtr_acco_numb like \'%" + pann + "\' and vtr_res4 = \'" + rrn + "\'")
        else:
            c.execute("select vtr_proc_date,vtr_acqu_refe_numb from visa_transaction_record where vtr_sour_amou = \'" + amountt + "\' and vtr_acco_numb like \'%" + pann + "\' and vtr_res4 = \'" + rrn + "\'")
    elif institution == 'CUP':
        c.execute("select utr_proc_date,utr_acqu_refe_numb from unp_transaction_record where utr_tran_amou_f004 = \'" + amountt + "\' and utr_acco_numb_f002 like \'%" + pann + "\' and (utr_retr_ref_numb_f037 = \'" + rrn + "\' or utr_res4 = \'" + rrn + "\')")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def process_tran_transaction_record(c, connection, rrn, pann, amountt,hst):
    rrn = str(rrn)
    pann = str(pann)
    amountt = str(amountt)
    c.execute("update svb_transaction_record set itr_vali_stat='VALI' where itr_tran_amou = \'" + amountt + "\' and itr_acco_numb like \'%" + pann + "\' and (itr_retr_refe_numb = \'" + rrn + "\' or itr_icc_issu_auth_data = \'" + rrn + "\')")
    connection.commit()


def check_clearing_detail(c,rrn,pann,amountt,hst):
    rrn = str(rrn)
    pann = str(pann)
    amountt = str(amountt)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select cld_set_date,cld_acqu_refe_numb from clearing_details where cld_tran_amou = \'" + amountt + "\' and cld_card_numb like \'%" + pann + "\' and cld_rrn_numb = \'" + rrn + "\' union "
                                                                                                                                                                                                  "select cld_set_date,cld_acqu_refe_numb from clearing_details_hst where cld_tran_amou = \'" + amountt + "\' and cld_card_numb like \'%" + pann + "\' and cld_rrn_numb = \'" + rrn + "\'")
    else:
        c.execute("select cld_set_date,cld_acqu_refe_numb from clearing_details where cld_tran_amou = \'" + amountt + "\' and cld_card_numb like \'%" + pann + "\' and cld_rrn_numb = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def check_remittance_transaction(c,rrn,pann,amountt):
    rrn = str(rrn)
    pann = str(pann)
    amountt = str(amountt)
    c.execute("select rtr_rem_code from remittance_transactions where (rtr_amou_tran = \'" + amountt + "\' or rtr_orig_amou = \'" + amountt + "\') and rtr_acc_numb like \'%" + pann + "\' and rtr_arch_refe = \'" + rrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def check_remittance(c,rem_code):
    rem_code = str(rem_code)
    c.execute(
        "select rem_code,rem_stat, rem_term_numb, rem_merc_iden,rem_ban_code,rem_star_date from remittance where rem_code = \'" + rem_code + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def update_remittance(c,connection,rem_code):
    rem_code = str(rem_code)
    c.execute("update remittance set rem_stat='Y' where rem_code = \'" + rem_code + "\'")
    connection.commit()


def split(word):
    return [char for char in word]


def get_date_format(date):
    if date == None:
        pass
    else:
        date = str(date)
        date = date.replace('-', '')
        dateall = split(date)
        dd = dateall[6] + dateall[7]
        mm = dateall[4] + dateall[5]
        yy = dateall[0] + dateall[1] + dateall[2] + dateall[3]
        datee = dd + '/' + mm + '/' + yy
        return datee

def get_date_format_for_reject(date):
    if date == None:
        pass
    else:
        date = str(date)
        date = date.replace('-', '')
        dateall = split(date)
        dd = dateall[6] + dateall[7]
        mm = dateall[4] + dateall[5]
        yy = dateall[0] + dateall[1] + dateall[2] + dateall[3]
        datee = yy + '/' + mm + '/' + dd
        return datee


def save_data(mycursor,FILE_NAME, INSTITUTION, BANK, RRN, PAN, AMOUNT, AQRRN = None, APPROVED_ON = None, INCLUDED_IN_OUTGOING  = None, CLEARING_DATE = None, FEEDBACK = None, ADDITIONAL_FEEDBACK = None, REQUEST_DATE = None, STATUS_CLOSED = None, STATUS = None):
    if INCLUDED_IN_OUTGOING == '' or INCLUDED_IN_OUTGOING == None:
       INCLUDED_IN_OUTGOING = None
    if CLEARING_DATE == '' or CLEARING_DATE == None:
       CLEARING_DATE = None
    if STATUS_CLOSED == '' or STATUS_CLOSED == None:
       STATUS_CLOSED = None
    if APPROVED_ON == '' or APPROVED_ON == None:
       APPROVED_ON = None
    if REQUEST_DATE == '' or REQUEST_DATE == None:
       REQUEST_DATE = None
    sql = "INSERT INTO `settle_all` (`FILE_NAME`, `INSTITUTION`, `BANK`, `RRN`, `PAN`, `AMOUNT`, `AQRRN`, `APPROVED_ON`, `INCLUDED_IN_OUTGOING`, `CLEARING_DATE`, `FEEDBACK`, `ADDITIONAL_FEEDBACK`, `REQUEST_DATE`, `STATUS_CLOSED`, `STATUS`) " \
          "VALUES (%s, %s,%s, %s,%s, %s,%s, %s,%s, %s,%s, %s,%s, %s,%s)"
    val = (FILE_NAME, INSTITUTION, BANK, RRN, PAN, AMOUNT, AQRRN, APPROVED_ON, INCLUDED_IN_OUTGOING, CLEARING_DATE, FEEDBACK, ADDITIONAL_FEEDBACK, REQUEST_DATE, STATUS_CLOSED, STATUS)
    mycursor.execute(sql, val)


def check_existance(mycursor,rrn,pann,amountt):
    rrn = str(rrn)
    pann = str(pann)
    amountt = str(amountt)
    mycursor.execute("select * from settle.settle_all where PAN like \'%" + pann + "\' and AMOUNT = \'" + amountt + "\' and RRN = \'" + rrn + "\'")
    while True:
        row = mycursor.fetchone()
        if row is None:
            break
        return row


def update_log_data(mycursor, RRN, PAN, AMOUNT, AQRRN = None, INCLUDED_IN_OUTGOING  = None, CLEARING_DATE = None,FEEDBACK = None, STATUS_CLOSED = None, STATUS = None):
    if INCLUDED_IN_OUTGOING == '' or INCLUDED_IN_OUTGOING == None:
       INCLUDED_IN_OUTGOING = ""
    if CLEARING_DATE == '' or CLEARING_DATE == None:
       CLEARING_DATE = ""
    if STATUS_CLOSED == '' or STATUS_CLOSED == None:
       STATUS_CLOSED = ""
    if AQRRN == '' or AQRRN == None:
       AQRRN = ""
    sql = "UPDATE `settle_all` SET `AQRRN`= \'" + AQRRN + "\',`INCLUDED_IN_OUTGOING`= \'" + INCLUDED_IN_OUTGOING + "\',`CLEARING_DATE`=\'" + CLEARING_DATE + "\',`FEEDBACK`=\'" + FEEDBACK + "\', `STATUS_CLOSED`=\'" + STATUS_CLOSED + "\',`STATUS`=\'" + STATUS + "\' WHERE `RRN`= \'" + RRN + "\' and `PAN` like \'%" + PAN + "\' and `AMOUNT` = \'" + AMOUNT + "\'"
    mycursor.execute(sql)




def get_ipm_itr_iou_code(c,aqrrn,hst):
    aqrrn = str(aqrrn)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select itr_iou_code from ipm_transaction_record where itr_inco_outg = 'OUTG' and itr_acqu_refe_numb = \'" + aqrrn + "\' union "
                                                                                                                                       "select itr_iou_code from ipm_transaction_record_hst where itr_inco_outg = 'OUTG' and itr_acqu_refe_numb = \'" + aqrrn + "\'")
    else:
        c.execute("select itr_iou_code from ipm_transaction_record where itr_inco_outg = 'OUTG' and itr_acqu_refe_numb = \'" + aqrrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row[0]

def get_ipm_itr_rre_code(c,aqrrn,hst):
    aqrrn = str(aqrrn)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select itr_rre_code from ipm_transaction_record where itr_inco_outg = 'INCO' and itr_acqu_refe_numb = \'" + aqrrn + "\' union "
                                                                                                                                       "select itr_rre_code from ipm_transaction_record_hst where itr_inco_outg = 'INCO' and itr_acqu_refe_numb = \'" + aqrrn + "\'")
    else:
        c.execute("select itr_rre_code from ipm_transaction_record where itr_inco_outg = 'INCO' and itr_acqu_refe_numb = \'" + aqrrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row[0]

def get_ipm_itr_iin_code(c,aqrrn,hst):
    aqrrn = str(aqrrn)
    hst = str(hst)
    if hst == 'Yesforhst':
        c.execute("select itr_iin_code from ipm_transaction_record where itr_inco_outg = 'INCO' and itr_acqu_refe_numb = \'" + aqrrn + "\' union "
                                                                                                                                       "select itr_iin_code from ipm_transaction_record_hst where itr_inco_outg = 'INCO' and itr_acqu_refe_numb = \'" + aqrrn + "\'")
    else:
        c.execute("select itr_iin_code from ipm_transaction_record where itr_inco_outg = 'INCO' and itr_acqu_refe_numb = \'" + aqrrn + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row[0]

def get_ipm_outgoing_data(c,iou_code):
    iou_code = str(iou_code)
    c.execute("select iou_proc_date from ipm_outgoing where iou_code = \'" + iou_code + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row

def get_ipm_outgoing_data_count(c,iou_code):
    iou_code = str(iou_code)
    c.execute("select count(*) from ipm_outgoing where iou_code = \'" + iou_code + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row[0]

def get_ipm_rejected_data(c,rre_code):
    rre_code = str(rre_code)
    c.execute("select rre_0280_file_refe_date  from ipm_rejected_record where rre_code = \'" + rre_code + "\'")
    # rre_proc_date
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row
def get_ipm_rejected_data_count(c,rre_code):
    rre_code = str(rre_code)
    c.execute("select count(*) rre_proc_date from ipm_rejected_record where rre_code = \'" + rre_code + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row[0]

def get_ipm_incoming_data(c,iin_code):
    iin_code = str(iin_code)
    c.execute("select iin_proc_date from ipm_incoming where iin_code = \'" + iin_code + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row


def get_ipm_incoming_data_by_file_id(c,iin_0105_file_sequ_numb):
    iin_0105_file_sequ_numb = str(iin_0105_file_sequ_numb)
    c.execute("select iin_proc_date,iin_0105_file_sequ_numb from ipm_incoming where iin_0105_file_sequ_numb = \'" + iin_0105_file_sequ_numb + "\'")
    while True:
        row = c.fetchone()
        if row is None:
            break
        return row



def main(atm,save,hst,fix,institution,output_file_name,bank, path, rrn_column, pan_column, amount_column):
    rem_code_list = []
    try:
        create_oracle_connection_result = create_oracle_connection()
        pool = create_oracle_connection_result[2]
        connection = create_oracle_connection_result[1]
        c = create_oracle_connection_result[0]
    except:
        print('Oracle Not Connected!!')
    if save == 'Yes':
        try:
            create_mysql_connection_result = create_mysql_connection()
            connection_pool = create_mysql_connection_result[2]
            connection_object = create_mysql_connection_result[1]
            mycursor = create_mysql_connection_result[0]
        except:
            print('Mysql Not Connected')
    try:
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(0)

        sheet.cell_value(0, 0)
        wbb = Workbook()
        result = wbb.add_sheet('Result')
        result_error = wbb.add_sheet('Rejected')
        settlement_not_done = wbb.add_sheet('Settlement Not Done')
        saved_requests = wbb.add_sheet('Requests Already Saved')
    except:
        print('Excel Read Error!')

    result.write(0, 0, "Bank")
    result.write(0, 1, "RRN")
    result.write(0, 2, "PAN")
    result.write(0, 3, "Amount")
    result.write(0, 4, "AqRRN")
    result.write(0, 5, "Approved On")
    result.write(0, 6, "Included in outgoing")
    result.write(0, 7, "Clearing date")
    result.write(0, 8, "Feedback")
    result.write(0, 9, "Additional Feedback")
    result.write(0, 10, "Request Date")
    result.write(0, 11, "Status Closed")
    result.write(0, 12, "Status")

    result_error.write(0, 0, "RRN")
    result_error.write(0, 1, "PAN")
    result_error.write(0, 2, "Amount")
    result_error.write(0, 3, "Reason")
    result_error.write(0, 4, "Additional Reason")

    settlement_not_done.write(0, 0, "RRN")
    settlement_not_done.write(0, 1, "PAN")
    settlement_not_done.write(0, 2, "Amount")
    settlement_not_done.write(0, 3, "Merchant ID")
    settlement_not_done.write(0, 4, "Rem Code")
    settlement_not_done.write(0, 5, "Terminal Numb")
    settlement_not_done.write(0, 6, "Bank Code")

    saved_requests.write(0, 0, "FILE NAME")
    saved_requests.write(0, 1, "INSTITUTION")
    saved_requests.write(0, 2, "BANK")
    saved_requests.write(0, 3, "RRN")
    saved_requests.write(0, 4, "PAN")
    saved_requests.write(0, 5, "AMOUNT")
    saved_requests.write(0, 6, "AQRRN")
    saved_requests.write(0, 7, "APPROVED ON")
    saved_requests.write(0, 8, "INCLUDED IN OUTGOING")
    saved_requests.write(0, 9, "CLEARING DATE")
    saved_requests.write(0, 10, "FEEDBACK")
    saved_requests.write(0, 11, "ADDITIONAL FEEDBACK")
    saved_requests.write(0, 12, "REQUEST DATE")
    saved_requests.write(0, 13, "STATUS CLOSED")
    saved_requests.write(0, 14, "STATUS")


    j, k, l, m, n= 1, 1, 1, 1, 1
    for i in range(sheet.nrows):
        bank_prev = bank
        refnum = str(sheet.cell_value(i, rrn_column)).strip()
        pann = str(sheet.cell_value(i, pan_column)).strip()
        amountt = str(sheet.cell_value(i, amount_column)).strip()
        if '.0' in refnum:
            refnum = str(int(float(refnum)))
        if len(refnum) < 12 and len(refnum) > 4:
            refnum = refnum.zfill(12)
        elif len(refnum) < 4:
            refnum = None
        if pann == 'Null':
            get_pann_result = get_pan(c,refnum, amountt,hst)
            if get_pann_result != None:
                pann = str(get_pann_result[0])
            elif get_pann_result == None:
                amountt = str(float(str(amountt)) + 0.01)
                get_pann_result = get_pan(c, refnum, amountt,hst)
                if get_pann_result != None:
                    pann = str(get_pann_result[0])
        if pann == 'Null':
            get_amount_and_pan_result = get_amount_and_pan(c,refnum,hst)
            if get_amount_and_pan_result != None:
                # result_error.write(j, 4, "With this RRN Amount = " + str(
                #     get_amount_and_pan_result[0]) + ", PAN = " + str(get_amount_and_pan_result[1]))
                amountt = str(get_amount_and_pan_result[0])
                pann = str(get_amount_and_pan_result[1])
                # j +=1
        if amountt == 'Null':
            get_amount_result = get_amount(c,refnum,pann,hst)
            if get_amount_result != None:
                amountt = str(get_amount_result[0])
        last_for_digit_pan = str(pann[-4:])
        pann = pann[:8] + "*****" + pann[12:]
        if refnum != None and len(pann) > 3:
            if save == 'Yes':
                check_existance_result = check_existance(mycursor,refnum, last_for_digit_pan, amountt)
            if save == 'Yes' and check_existance_result != None:
                saved_requests.write(m, 0, check_existance_result[1])
                saved_requests.write(m, 1, check_existance_result[2])
                saved_requests.write(m, 2, check_existance_result[3])
                saved_requests.write(m, 3, check_existance_result[4])
                saved_requests.write(m, 4, check_existance_result[5])
                saved_requests.write(m, 5, check_existance_result[6])
                saved_requests.write(m, 6, check_existance_result[7])
                if check_existance_result[8] != None:
                    saved_requests.write(m, 7, get_date_format(check_existance_result[8]))
                if check_existance_result[9] != None:
                    saved_requests.write(m, 8, get_date_format(check_existance_result[9]))
                if check_existance_result[10] != None:
                    saved_requests.write(m, 9, get_date_format(check_existance_result[10]))
                saved_requests.write(m, 10, check_existance_result[11])
                saved_requests.write(m, 11, check_existance_result[12])
                if check_existance_result[13] != None:
                    saved_requests.write(m, 12, get_date_format(check_existance_result[13]))
                if check_existance_result[14] != None:
                    saved_requests.write(m, 13, get_date_format(check_existance_result[14]))
                saved_requests.write(m, 14, check_existance_result[15])
                m += 1
            else:
                global authorization_result
                if atm != 'Yess':
                    authorization_result = check_authorization(c,refnum,last_for_digit_pan,amountt,hst)
                else:
                    authorization_result = check_authorization_for_atm(c,refnum,last_for_digit_pan,amountt,hst)
                acqu_refe_numb = ''
                tran_proc_date = ''
                if bank_prev == 'AUTO_SELECT':
                    if authorization_result == None or authorization_result[6] == '':
                        bank = 'BANK_NOT_FOUND'
                    else:
                        get_bank_result = get_bank(authorization_result[6])
                        bank = get_bank_result
                print(authorization_result)
                if authorization_result == None:
                    pos_transaction_result = check_pos_transaction(c,refnum,last_for_digit_pan,amountt,hst)
                    check_amount_result = check_amount(c,refnum,last_for_digit_pan,hst)
                    if pos_transaction_result != None and atm != 'Yess':
                        result_error.write(j, 0, refnum)
                        result_error.write(j, 1, pann)
                        result_error.write(j, 2, amountt)
                        result_error.write(j, 3, "Not POS transaction")
                        j += 1

                    elif check_amount_result != None:
                        correct_amount = check_amount_result[2]
                        result_error.write(j, 0, refnum)
                        result_error.write(j, 1, pann)
                        result_error.write(j, 2, amountt)
                        result_error.write(j, 3, "Incorrect Amount")
                        result_error.write(j, 4, "With this RRN and PAN the Amount is => " + str(correct_amount))
                        j += 1
                    else:
                        result_error.write(j, 0, refnum)
                        result_error.write(j, 1, pann)
                        result_error.write(j, 2, amountt)
                        result_error.write(j, 3, "RRN, PAN, and Amount Mismatched")
                        j += 1

                elif authorization_result != None:
                    check_reversal_status_result = check_reversal_status(c,refnum,last_for_digit_pan,amountt,hst)
                    check_pre_aut_status_result = check_pre_aut_status(c,refnum,last_for_digit_pan,amountt,hst)
                    if check_reversal_status_result != None:
                        result_error.write(j, 0, refnum)
                        result_error.write(j, 1, pann)
                        result_error.write(j, 2, amountt)
                        result_error.write(j, 3, "Transaction Reversed")
                        result_error.write(j, 4, get_date_format(check_reversal_status_result[0]))
                        j += 1
                    elif check_pre_aut_status_result != None:
                        result_error.write(j, 0, refnum)
                        result_error.write(j, 1, pann)
                        result_error.write(j, 2, amountt)
                        result_error.write(j, 3, "Pre_aut transaction, completion must be done")
                        result_error.write(j, 4, get_date_format(check_pre_aut_status_result[0]))
                        j += 1
                    else:
                        rrn = authorization_result[0]
                        response_code = authorization_result[1]
                        amount = authorization_result[2]
                        pan = authorization_result[3]
                        inte_code = authorization_result[4]  # Pos or atm
                        authorized_date = authorization_result[5]
                        if rrn != None and amount != None and pan != None and inte_code != None and authorized_date != None:
                            if response_code != '000':
                                result_error.write(j, 0, refnum)
                                result_error.write(j, 1, pann)
                                result_error.write(j, 2, amount)
                                result_error.write(j, 3, "Not Approved Transaction")
                                result_error.write(j, 4, "Response Code = " + str(response_code))
                                j += 1
                                k += 1
                            elif response_code == '000':
                                result.write(l, 10, today_date)
                                result.write(l, 0, bank)
                                result.write(l, 1, refnum)
                                result.write(l, 2, pann)
                                result.write(l, 3, amountt)
                                result.write(l, 5, get_date_format(authorized_date))
                                tran_transaction_result = check_tran_transaction_record(c,institution,refnum, last_for_digit_pan, amountt,hst)
                                clearing_result = check_clearing_detail(c,refnum,last_for_digit_pan,amount,hst)
                                if tran_transaction_result != None:
                                    tran_proc_date = tran_transaction_result[0]
                                    acqu_refe_numb = str(tran_transaction_result[1])
                                    result.write(l, 4, acqu_refe_numb)
                                    if clearing_result != None:
                                        cld_set_date = clearing_result[0]
                                        result.write(l, 7, get_date_format(cld_set_date))
                                    if tran_proc_date != None:
                                        result.write(l, 6, get_date_format(tran_proc_date))
                                        result.write(l, 8, "Pre_Generated")
                                        result.write(l, 12, "Closed")
                                        # process_tran_transaction_record(c, connection, refnum, last_for_digit_pan,amountt,hst)


                                        if institution == 'MasterCard':
                                            rejected_data = ''
                                            incoming_data = ''
                                            outgoing_data = ''
                                            outgoing_date = ''
                                            outgoing_file_id = ''
                                            rejected_code = get_ipm_itr_rre_code(c,acqu_refe_numb,hst)
                                            incoming_code = get_ipm_itr_iin_code(c,acqu_refe_numb,hst)
                                            outgoing_code = get_ipm_itr_iou_code(c,acqu_refe_numb,hst)
                                            # if outgoing_code[0] == None:
                                            #     print('No outgoing code')
                                            # elif outgoing_code != None:
                                            #     print('Outgoing code found')
                                            # if rejected_code[0]==None:
                                            #     print('No rejected code')
                                            # elif rejected_code[0] != None:
                                            #     print('Rejected code found')
                                            # if incoming_code[0] == None:
                                            #     print('No incoming code')
                                            # elif incoming_code[0] != None:
                                            #     print('Incoming code found')

                                            if incoming_code != None and rejected_code == None:
                                                incoming_data = get_ipm_incoming_data(c, incoming_code)
                                                inc_date = get_date_format(incoming_data[0])
                                                print('Closed')
                                                result.write(l, 9, 'Closed_outgoing_'+inc_date)
                                            elif incoming_code != None and rejected_code != None and outgoing_code != None:
                                                outgoing_data_count = get_ipm_outgoing_data_count(c, outgoing_code)
                                                rejected_data_count = get_ipm_rejected_data_count(c,rejected_code)

                                                if outgoing_data_count > 1:
                                                    print("First Present Needs To Be Checked")
                                                if rejected_data_count>1:
                                                    print(rejected_data_count)
                                                outgoing_data = get_ipm_outgoing_data(c, outgoing_code)
                                                rejected_data = get_ipm_rejected_data(c, rejected_code)
                                                incoming_data = get_ipm_incoming_data(c, incoming_code)
                                                out_date = get_date_format_for_reject(outgoing_data[0])
                                                rej_date = get_date_format_for_reject(rejected_data[0])
                                                inc_date = get_date_format_for_reject(incoming_data[0])
                                                if out_date > rej_date:
                                                    print('Closed')
                                                    result.write(l, 9, 'Closed_File_ref_date_'+ rej_date+', outgoing_'+out_date)
                                                elif out_date < rej_date:
                                                    print('Rejected')
                                                    result.write(l, 9, 'Rejected_'+ rej_date+', outgoing_'+out_date)
                                                    print(out_date,rej_date)
                                                else:
                                                    print('Still Rejected')
                                                    result.write(l, 9, 'Still Rejected_'+ rej_date+', outgoing_'+out_date)
                                                # print('Date should be compared')
                                            elif incoming_code == None and rejected_code == None:
                                                print('Integration not done or no incoming')
                                                result.write(l, 9, 'Integration_not_done_or_no_incomming')
                                            else:
                                                print('undetected')





                                        result.write(l, 11, today_date)
                                        if save == 'Yes':
                                            save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                      str(acqu_refe_numb), authorized_date, tran_proc_date,
                                                      cld_set_date, "Pre_Generated", "",
                                                      today, today, "Closed")
                                    elif tran_proc_date == None:
                                        result.write(l, 8, "Not_Processed_Yet_After_EOD)")
                                        result.write(l, 12, "Pending")
                                        if save == 'Yes':
                                            save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                      str(acqu_refe_numb), authorized_date, tran_proc_date,
                                                      cld_set_date, "Not_Processed_Yet_After_EOD", "",
                                                      today, today, "Pending")
                                    else:
                                        result.write(l, 8, "Uncaught_Tran_transaction_record_table_error")
                                        result.write(l, 12, "Pending")
                                        if save == 'Yes':
                                            save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                      str(acqu_refe_numb), authorized_date, tran_proc_date,
                                                      cld_set_date, "Uncaught_Tran_transaction_record_table_error", "",
                                                      today, today, "Pending")

                                elif tran_transaction_result == None:
                                    result_check_ipm_transaction_type = 'Y'
                                    if institution == 'MasterCard':
                                        result_check_ipm_transaction_type = check_ipm_transaction_type(c, refnum, last_for_digit_pan, amountt, hst)
                                        if(result_check_ipm_transaction_type == None):
                                            result.write(l, 9, "No_outgoing_for_MDS")
                                    remittance_transaction_result = check_remittance_transaction(c,refnum, last_for_digit_pan,amountt)
                                    if clearing_result != None and institution != 'Local_Transaction':
                                        cld_set_date = clearing_result[0]
                                        cld_acqu_refe_numb = str(clearing_result[1])
                                        result.write(l, 4, cld_acqu_refe_numb)
                                        result.write(l, 7, get_date_format(cld_set_date))
                                        if result_check_ipm_transaction_type == None:
                                            result.write(l, 8,"MDS_In_Clearing")
                                            result.write(l, 11, today_date)
                                            result.write(l, 12, "Closed")
                                        else:
                                            result.write(l, 8, "IN_CL_NOT_OUT_S2M")
                                            result.write(l, 12, "Pending")
                                        if save == 'Yes':
                                            if result_check_ipm_transaction_type == None:
                                                save_data(mycursor, output_file_name, institution, bank, refnum, pann,
                                                          amountt,
                                                          str(cld_acqu_refe_numb), authorized_date, tran_proc_date,
                                                          cld_set_date, "MDS_In_Clearing", "No_outgoing_for_MDS",
                                                          today, today, "Closed")
                                            else:
                                                save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                          str(cld_acqu_refe_numb), authorized_date, tran_proc_date,
                                                          cld_set_date, "IN_CL_NOT_OUT_S2M", "",
                                                          today, "", "Pending")
                                    elif clearing_result != None and institution == 'Local_Transaction':
                                        cld_set_date = clearing_result[0]
                                        cld_acqu_refe_numb = str(clearing_result[1])
                                        result.write(l, 4, cld_acqu_refe_numb)
                                        result.write(l, 7, get_date_format(cld_set_date))
                                        result.write(l, 8, "Pre_Included_In_Clearing")
                                        result.write(l, 12, "Closed")
                                        result.write(l, 11, today_date)
                                        if save == 'Yes':
                                            save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                      str(acqu_refe_numb), authorized_date, tran_proc_date,
                                                      cld_set_date, "Pre_Included_In_Clearing", "",
                                                      today, today, "Closed")
                                    else:
                                        if remittance_transaction_result != None:
                                            rem_code = remittance_transaction_result[0]
                                            remittance_result = check_remittance(c,rem_code)
                                            rem_date = remittance_result[5]
                                            if remittance_result != None:
                                                rem_stat = remittance_result[1]
                                                if rem_stat == None:
                                                    rem_term_numb = remittance_result[2]
                                                    rem_merc_iden = remittance_result[3]
                                                    rem_ban_code = remittance_result[4]
                                                    if rem_term_numb != None and rem_merc_iden != None and rem_ban_code != None:
                                                        settlement_not_done.write(n, 0, refnum)
                                                        settlement_not_done.write(n, 1, pann)
                                                        settlement_not_done.write(n, 2, amountt)
                                                        settlement_not_done.write(n, 3, rem_merc_iden)
                                                        settlement_not_done.write(n, 4, rem_code)
                                                        settlement_not_done.write(n, 5, rem_term_numb)
                                                        settlement_not_done.write(n, 6, rem_ban_code)
                                                        result.write(l, 8, "SETTLE_NOT_DONE_EOD")
                                                        result.write(l, 12, "Pending")
                                                        if save == 'Yes':
                                                            save_data(mycursor,output_file_name, institution, bank, refnum, pann,
                                                                      amountt,
                                                                      (acqu_refe_numb), authorized_date, tran_proc_date,
                                                                      cld_set_date,
                                                                      "SETTLE_NOT_DONE_EOD", "",
                                                                      today, "", "Pending")
                                                        n += 1
                                                elif rem_stat == 'N':
                                                    rem_code_list.append(rem_code)
                                                    if (fix == 'Yesforfix'):
                                                        update_remittance(c,connection,rem_code)
                                                        result.write(l, 8, "REM_STAT_CHANGED_EOD")
                                                    result.write(l, 8, "REM_STAT_MUST_BE_CHANGED_EOD")
                                                    result.write(l, 9, rem_code)
                                                    result.write(l, 12, "Pending")
                                                    if save == 'Yes':
                                                        save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                                  (acqu_refe_numb), authorized_date, tran_proc_date,
                                                                  cld_set_date, "REM_STAT_MUST_BE_CHANGED_EOD", "",
                                                                  today, "", "Pending")
                                                elif rem_stat == 'Y':
                                                    if rem_code in rem_code_list:
                                                        result.write(l, 8, "REM_STAT_CHANGED_EOD")
                                                        result.write(l, 9, rem_code)
                                                        result.write(l, 12, "Pending")
                                                        if save == 'Yes':
                                                            save_data(mycursor, output_file_name,
                                                                      institution, bank, refnum, pann, amountt,
                                                                      (acqu_refe_numb), authorized_date, tran_proc_date,
                                                                      cld_set_date, "REM_STAT_CHANGED_EOD", "",
                                                                      today, "", "Pending")
                                                    else:
                                                        result.write(l, 8, "Y_S2M")
                                                        rem_datee = "Settlement Date = >" + str(rem_date)
                                                        result.write(l, 9, rem_datee)
                                                        result.write(l, 12, "Pending")
                                                        if save == 'Yes':
                                                            save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                                      (acqu_refe_numb), authorized_date, tran_proc_date,
                                                                      "", "Y_S2M",rem_datee,
                                                                      today, "", "Pending")
                                                    m += 1
                                                else:
                                                    result.write(l, 8, "UNCAUGHT_REM_TAB_ERROR")
                                                    result.write(l, 12, "Pending")
                                                    if save == 'Yes':
                                                        save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                                  (acqu_refe_numb), authorized_date, tran_proc_date,
                                                                  cld_set_date, "UNCAUGHT_REM_TAB_ERROR", "",
                                                                  today, "", "Pending")
                                            elif remittance_result == None:
                                                result.write(l, 8, "REM_DATA_NOT_FOUND_S2M")
                                                result.write(l, 12, "Pending")
                                                if save == 'Yes':
                                                    save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                              (acqu_refe_numb), authorized_date, tran_proc_date,
                                                              cld_set_date, "REM_DATA_NOT_FOUND_S2M", "",
                                                              today, "", "Pending")
                                        elif remittance_transaction_result == None:
                                            result.write(l, 8, "REM_C_NOT_FOUND_S2M")
                                            result.write(l, 12, "Pending")
                                            if save == 'Yes':
                                                save_data(mycursor,output_file_name, institution, bank, refnum, pann, amountt,
                                                          (acqu_refe_numb), authorized_date, tran_proc_date,
                                                          cld_set_date, "REM_C_NOT_FOUND_S2M", "",
                                                          today, "", "Pending")
                                l += 1
                        else:
                            result_error.write(j, 0, refnum)
                            result_error.write(j, 1, pann)
                            result_error.write(j, 2, amount)
                            result_error.write(j, 3, "Uncaught Authorization Error")
                            j += 1
        authorized_date, tran_proc_date, cld_set_date, acqu_refe_numb, refnum, pann, amountt, acqu_refe_numb = '','','', '', '', '', '', ''
        bank = bank_prev
    output_path = os.getcwd()
    dir = os.getcwd() + '/' + str(output_file_name)

    if os.path.exists(dir):
        shutil.rmtree(dir)
    os.mkdir(os.path.join(output_path, str(output_file_name)))

    wbb.save(os.path.join(os.path.join(output_path, str(output_file_name)), str("RESULT."+output_file_name + '.xls')))
    close_oracle_connection(pool,connection)
    if save == 'Yes':
        close_mysql_connection(connection_object,mycursor)

def update_log():
    hst = 'No'
    try:
        create_oracle_connection_result = create_oracle_connection()
        pool = create_oracle_connection_result[2]
        connection = create_oracle_connection_result[1]
        c = create_oracle_connection_result[0]
    except:
        print('Oracle Not Connected!!')
    create_mysql_connection_result = create_mysql_connection()
    connection_object = create_mysql_connection_result[1]
    mycursor = create_mysql_connection_result[0]
    print('Update Request ...')
    mycursor.execute("select RRN, PAN, AMOUNT, INSTITUTION from settle.settle_all where Status = 'Pending'")
    while True:
        for row in mycursor.fetchall():
            refnum = str(row[0]).strip()
            pann = str(row[1]).strip()
            amountt = str(row[2]).strip()
            institution = str(row[3]).strip()
            last_for_digit_pan = str(pann[-4:])
            if refnum != None and len(pann) > 3:
                authorization_result = check_authorization(c, refnum, last_for_digit_pan, amountt,hst)
                print(authorization_result)
            if authorization_result != None:
                check_reversal_status_result = check_reversal_status(c, refnum, last_for_digit_pan, amountt,hst)
                if check_reversal_status_result != None:
                    print('Reversed!!')
                else:
                    rrn = authorization_result[0]
                    response_code = authorization_result[1]
                    amount = authorization_result[2]
                    pan = authorization_result[3]
                    inte_code = authorization_result[4]  # Pos or atm
                    authorized_date = authorization_result[5]
                    if rrn != None and amount != None and pan != None and inte_code != None and authorized_date != None:
                        if response_code != '000':
                            print('Not Approved Transaction')
                        elif response_code == '000':
                            tran_transaction_result = check_tran_transaction_record(c, institution, refnum,last_for_digit_pan, amountt,hst)
                            clearing_result = check_clearing_detail(c, refnum, last_for_digit_pan, amount,hst)
                            if institution == 'Local_Transaction':
                                if clearing_result != None:
                                    cld_set_date = clearing_result[0]
                                    print('Update Found!!' + str(refnum))
                                    update_log_data(mycursor, str(refnum), str(last_for_digit_pan), str(amountt), str(acqu_refe_numb),str(tran_proc_date),str(cld_set_date), "Generated_"+get_date_format(today), str(today), "Closed")
                                    connection_object.commit()
                            if tran_transaction_result != None:
                                tran_proc_date = tran_transaction_result[0]
                                acqu_refe_numb = str(tran_transaction_result[1])
                                if clearing_result != None:
                                    cld_set_date = clearing_result[0]
                                if tran_proc_date != None or clearing_result != None:
                                    print('Update Found!!' + str(refnum))
                                    update_log_data(mycursor, str(refnum), str(last_for_digit_pan), str(amountt), str(acqu_refe_numb),str(tran_proc_date),str(cld_set_date), "Generated_"+get_date_format(today), str(today), "Closed")
                                    connection_object.commit()
            authorized_date, tran_proc_date, cld_set_date, acqu_refe_numb, refnum, pann, amountt, acqu_refe_numb = '', '', '', '', '', '', '', ''
        else:
            break

    close_oracle_connection(pool, connection)
    close_mysql_connection(connection_object, mycursor)