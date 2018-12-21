import sys, traceback, os

def get_num_xlsx(name):
    return int(name[-7:-5])

def get_num(name):
    return int(name[-2:])

def get_num_s(name):
    print(name[-1:])
    return int(name[-1:])

def get_num_sxlsx(name):
    return int(name[-6:-5])

def get_new_invoice_num(folder):
    try:
        invoice_list    = os.listdir(folder)
        max_num = 0
        for x in invoice_list:
            if 'invoice' not in x:
                continue
            if len(x) == 8 or len(x) == 13:
                if x.endswith('.xlsx'):
                    num = get_num_sxlsx(x)
                    if num > max_num:
                        max_num = num
                else:
                    num = get_num_s(x)
                    if num > max_num:
                        max_num = num

            elif x.endswith('.xlsx'):
                num = get_num_xlsx(x)
                if num > max_num:
                    max_num = num
            else:
                num = get_num(x)
                if num > max_num:
                    max_num = num
    except:
        print('[!] Failed To Get Invoice Number')
        traceback.print_exc(file=sys.stdout)
        sys.exit(1)
    return str(max_num+1)
