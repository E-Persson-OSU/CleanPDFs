#script to organize and name PDFs in PDF folder on my H drive and send me a report via email


#imports
from os               import listdir, rename
from os.path          import getctime
from time             import localtime, strftime
from ntpath           import basename
from smtplib          import SMTP_SSL
from email.mime.text  import MIMEText
from configparser     import RawConfigParser

#config loading
config = RawConfigParser()
config.read('cleanpdf.ini')
login, pw, rcp = config['login'], config['password'], config['recipient']

#begin functions

#verifies file is a pdf
def ispdf(file):
    return file.lower().endswith('.pdf')

#Takes file and daily identifier, checks proper directory for file and renames it and moves it to that directory with new file name
def renameandfile(file,di):

    #name file
    ctf = strftime("%Y_%m_%d_%H_%M_%S", localtime(getctime(file)))        # creation time formatted to YEAR_MONTH_DAY_HOUR_MINUTE_SECOND 
    ctf = ctf + '_' + str(di) + '.pdf'                                    # Adding unique daily identifier to prevent same name files and file extension

    #Determine correct folder for file
    bn = basename(file)
    folder = 'Misc\\'
    if bn.lower().startswith('ps'):
        folder = 'Packing Slip\\'
    elif bn.lower().startswith('rs'):
        folder = 'Return Slips\\'
    elif bn.lower().startswith('qu'):
        folder = 'Quotes\\'

    #rename file
    rename(file, pfp + folder + ctf)

    #return action string
    return "Renamed and moved " + file + " to " + pfp + folder + ctf + "\n"


#TODO make python file for sending emails

#makes report from action list, saves it to a text document in separate folder
def mkreport(actionlist):
    fnw = 'H:\\Scripts\\PDFCleanupReports\\' + strftime('%d_%m_%Y') + '.txt'
    with open(fnw ,'w+') as fw:
        for a in actionlist:
            fw.write(a)
        fw.write("\nEnd report.")
    #returns file name for later use
    return fnw

def sendreport(filename):
    date = strftime("%a, %m %d" ,localtime())
    with open(filename) as re:
        msg = MIMEText(re.read())
    
    msg['Subject'] = "PDF Cleanup on " + date
    msg['From'] = login
    msg['To'] = rcp

    s = SMTP_SSL('smtp.gmail.com',465,timeout=10)
    try:
        s.login(login,pw)
        s.sendmail(msg['From'], msg['To'], msg.as_string())
    finally:
        s.quit()



#end functions

#----------BEGIN MAIN----------

#PDF Folder Directory
pfp = 'H:\\pdfs\\'
#get files in directory
pdflist = listdir(pfp)
#Daily identifier equal to n, where n is index of file in pdflist
di = 1
#action list to record actions taken during script
al = []
for f in pdflist:
    fn = pfp + f
    if ispdf(fn):
        al.append(renameandfile(fn,di))
        di += 1
msgfile = mkreport(al)
sendreport(msgfile)
#-----------END MAIN-----------