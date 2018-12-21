#!/usr/bin/env python3
import argparse
import emailer
import getpass
import constants
import tools
import os
import time
import xml2xlsx
from framework import ImageReader, InvoiceReader, InvoiceWriter

def getJobNum(s):
    left, right = s.index("#") + 2, s.find("=", s.index('#'))-1
    return s[left:right]

def main():
    start = time.time()

    parser = argparse.ArgumentParser(description="Automate Invoice")
    parser.add_argument("-f", "--file", type=str, required=True, metavar="", help="Image Path")
    parser.add_argument("-r", "--receiver", type=str, required=False, metavar="", help="Email receiver")
    parser.add_argument("-u", "--user", type=str, required=False, metavar="", help="Email sender")
    parser.add_argument("-p", "--psm", type=str, metavar="", help="Specify image psm")

    args = parser.parse_args()
    user = args.user
    receiver = args.receiver
    psm = args.psm
    file = args.file

    if user == None:
        user = constants.DEFAULT_SENDER

    if receiver == None:
        receiver = constants.DEFAULT_RECEIVER

    new_invoice_num = tools.get_new_invoice_num(constants.DANNY_FOLDER_PATH)

    save_path = os.path.join(constants.DANNY_FOLDER_PATH, "invoice" + new_invoice_num + ".xlsx")
 
    imageReader = ImageReader(file, psm)

    invoiceWriter = InvoiceWriter(constants.SAMPLE_FILE_PATH)
    invoiceWriter.writeInvoiceDateCreation(input("\nCreation Date: "))
    invoiceWriter.writeInvoiceNumber(new_invoice_num)

    password = getpass.getpass("Password For {}: ".format(user))
    email = emailer.IMAPEmailer(user, password, constants.YAHOO_IMAP_SERVER)
    print("\nRetrieving settlement file from email list\n")
    email.retrieveMostRecentFileWithExt(constants.DEFAULT_EXTENSION, constants.SETTLEMENT_FILE_PATH, constants.DEFAULT_PERSON)
    email.close()
    print("\nChanging extention from xml to xlsx\n")
    xml2xlsx.xml2xlsx(constants.SETTLEMENT_FILE_PATH, constants.RECENT_INVOICE_FILE_PATH)

    invoiceReader = InvoiceReader(constants.RECENT_INVOICE_FILE_PATH)

    jobs = imageReader.getSanitizeStringFromImage()

    ans = input("\nFound # Of Jobs: {} (y/n) ".format(len(jobs)))
    
    if ans.lower() != 'y':
        exit(1)

    row = constants.DEFAULT_ROW_START

    total = 0

    notfound = list()

    for job in jobs:
        info = invoiceReader.getInfoDictionary(getJobNum(job))
        if info['amt'] == None:
            info = invoiceReader.getInfoDictionary(input("\nFound job {}; enter correction: ".format((getJobNum(job)))))
        if info['amt'] == None:
            notfound.append(getJobNum(job))
            continue
        total += float(info['amt'])
        invoiceWriter.writeJob(row, info['date'], info['trackID'], info['job'], info['from'], info['to'], info['amt'])
        row += 2

    invoiceWriter.writeTotal(total, constants.DANNY_TOTAL_ROW)
    invoiceWriter.finalize(save_path)
    
    os.popen("open " + constants.RECENT_INVOICE_FILE_PATH)
    os.popen("open " + save_path)
    
    if len(notfound) > 0:
        print("\nJobs not in settlement {}:\n".format(len(notfound)))
        for job in notfound:
            print("\n{}\n".format(job))
    
    input("\nMake final adjustments; enter anything to send ")

    email = emailer.SMTPEmailer(constants.DEFAULT_SENDER, password, constants.YAHOO_SMTP_SERVER)
    email.sendattachment(os.path.basename(save_path), user, save_path)
    print("\nSent {} with subject {} to {}\n".format(save_path, os.path.basename(save_path), user))
    email.sendattachment(os.path.basename(save_path), receiver, save_path)
    print("\nSent {} with subject {} to {}\n".format(save_path, os.path.basename(save_path), receiver))
    email.close()

    print("\nFinished Job in ######### -> {:.2f} seconds\n".format(time.time()-start))
if __name__ == '__main__':
    main()

