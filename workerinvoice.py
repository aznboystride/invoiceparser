#!/usr/bin/env python3
import argparse
import getpass
import constants
import tools
import os

def main():
    
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
 
    imageReader = ImageReader(pathToImage, psm)

    invoiceWriter = InvoiceWriter(constants.SAMPLE_FILE_PATH)
    invoiceWriter.writeInvoiceDateCreation(input("\nCreation Date: "))
    invoiceWriter.writeInvoiceNumber(new_invoice_num)

    password = getpass.getpass("Password For {}: ".format(user))
    email = IMAPEmailer(user, password, constants.YAHOO_IMAP_SERVER)
    email.retrieveMostRecentFileWithExt(constants.DEFAULT_EXTENTION, constants.SETTLEMENT_FILE_PATH, constants.DEFAULT_PERSON)
    print("\nChange extension to xlsx\n")
    os.popen("open {}".format(constants.SETTLEMENT_FILE_PATH))
    input("\nEnter any key after changing extension\n")

    invoiceReader = InvoiceReader(constants.RECENT_INVOICE_FILE_PATH)

    jobs = imageReader.getSanitizeStringFromImage()

    ans = input("\nFound # Of Jobs: {} (y/n)".format(len(jobs)))
    
    if ans.lower() != 'y':
        exit(1)

    row = constants.DEFAULT_ROW_START

    total = 0

    for job in jobs:
        info = invoiceReader.getInfoDictionary(job)
         
    


if __name__ == '__main__':
    main()

