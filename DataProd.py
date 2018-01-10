import xmltodict
from datetime import datetime, timedelta
import win32com.client
import json
import os

THIS_FILE = os.path.abspath(__file__)
PROJECT_DIR = os.path.dirname(THIS_FILE)


def get_dataprod_status():
    config_file = os.path.join(PROJECT_DIR, 'configs.json')
    with open(config_file) as fh:
        configs = json.load(fh)
    with open(configs['status']) as fd:
        doc = xmltodict.parse(fd.read())
    process = doc['DataProductionStatus']['Status']['@Currently']
    processstarted = doc['DataProductionStatus']['Status']['@ProcessStartedWhen']
    runlastreported = doc['DataProductionStatus']['Status']['@RunLastReported']

    return {'process': process, 'process_started': processstarted, 'run_reported': runlastreported}


def get_hours_when_expire(process):
    return {
        'endofweek': 5,
        'endofmonth': 5,
        'standard': 2,
    }.get(process, 2)


def send_cdo_msg(recipients, subject, body, sent_from):
    conf = win32com.client.Dispatch("CDO.Configuration")
    flds = conf.Fields
    flds("http://schemas.microsoft.com/cdo/configuration/smtpserver").Value = "eulofmsmtp01.economistgroup.net"
    flds("http://schemas.microsoft.com/cdo/configuration/smtpserverport").Value = 25
    flds("http://schemas.microsoft.com/cdo/configuration/sendusing").Value = 2  # cdoSendUsingPort
    # Authentication and stuff
    flds('http://schemas.microsoft.com/cdo/configuration/smtpauthenticate').Value = 0  # No authentication
    # The following fields are only used if the previous authentication value is set to 1 or 2
    # flds('http://schemas.microsoft.com/cdo/configuration/smtpaccountname').Value = "user"
    # flds('http://schemas.microsoft.com/cdo/configuration/sendusername').Value = "elzbietawatroba@economist.com"
    flds('http://schemas.microsoft.com/cdo/configuration/smtpusessl').Value = False
    # flds('http://schemas.microsoft.com/cdo/configuration/sendpassword').Value = "password"
    flds.Update()
    msg = win32com.client.Dispatch("CDO.Message")
    msg.Configuration = conf
    msg.To = recipients
    msg.From = sent_from
    msg.Subject = subject
    msg.TextBody = body
    msg.Send()


def is_dataprod_running():
    config_file = os.path.join(PROJECT_DIR, 'configs.json')
    status = get_dataprod_status()
    lastrun_date = datetime.strptime(status['run_reported'], '%d %b %Y %H:%M:%S')
    process = status['process']
    h = get_hours_when_expire(process)
    expired_time = datetime.now() - timedelta(hours=h)
    if expired_time > lastrun_date:
        with open(config_file) as fh:
            configs = json.load(fh)
        body = configs['body'] + datetime.strftime(lastrun_date, '%d %b %Y %H:%M') + '.'
        send_cdo_msg(configs['recipient'], configs['subject'], body, configs['sender'])


if __name__ == "__main__":
    is_dataprod_running()
