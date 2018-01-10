import gspread
from oauth2client.service_account import ServiceAccountCredentials
from sets import Set
# authorize gspread
scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('IDP_keys.json', scope)

gc = gspread.authorize(credentials)


# open worksheet
spreadsheet = gc.open_by_key('1a4-t-LI_IwfcM02byKARwBcU5-nIEMzFy_ohy_rtp1Q')
# won't open...
    # resolution!: go to spreadsheet and share sheet w/ client_email in .json keys file

# get prompts and responses to design, technology, and research sheets
def DTRFreeResponses():
    for page_number in range(3, 6):
        sheet = spreadsheet.get_worksheet(page_number)
        for row in range(3, 6):
            prompt = sheet.acell('A%d' % row).value
            response = sheet.acell('B%d' % row).value
            print prompt + ': ' + response
            print
        print

def LIPFreeResponse():
    sheet = spreadsheet.worksheet('LIP')
    for row in range(5, 8):
        prompt = sheet.acell('A%d' % row).value
        response = sheet.acell('B%d' % row).value
        print prompt + ': ' + response
        print

def CollaborationFreeResponse():
    sheet = spreadsheet.worksheet('Collaboration')
    for row in range(4, 7):
        prompt = sheet.acell('A%d' % row).value
        response = sheet.acell('B%d' % row).value
        print prompt + ': ' + response
        print

def GrowthFreeResponse():
    sheet = spreadsheet.worksheet('Growth')
    for row in range(5, 8):
        prompt = sheet.acell('A%d' % row).value
        response = sheet.acell('B%d' % row).value
        print prompt + ': ' + response


def DTRProcess():
    # would be usefull to tally which cells are called on often b/w assessments
    # not using 11, 12, 13, 16, 20, 32
    not_agree_count = 0
    # skip_list = Set([11, 12, 13, 16, 20, 32])
    rows = [2, 3, 4, 5, 6, 7, 8, 9, 10, 14, 15, 17, 18, 19,
    21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34]
    sheet = spreadsheet.worksheet('DTR Process')
    # for row in range(2, 35):
    #     if row not in skip_list:
    for row in rows:
        prompt = sheet.acell('B%d' % row).value
        not_agree_dict = {'C': 'N/A or DID NOT DO THIS', 'D': 'STRONGLY DISAGREE', 'E': 'DISAGREE', 'F': 'NEITHER AGREE NOR DISAGREE'}
        for column in not_agree_dict.keys():
            val = sheet.acell('%c%d' % (column,row)).value
            if val: # if student answered anything other than agree or strongly agree, print the response and the propmt
                print not_agree_dict[column] + ": " + prompt
                not_agree_count += 1
    print not_agree_count

def ResearchProcess():
    # would be usefull to tally which cells are called on often b/w assessments
    # not using 2, 10, 11, 12, 16, 24, 25, 29
    not_agree_count = 0
    # skip_list = Set([10, 11, 12, 16, 24, 25, 29])
    rows = [3, 4, 5, 6, 7, 8, 9, 13, 14, 15, 17, 18, 19, 20, 21,
    22, 23, 26, 27, 28, 30, 31, 32, 33, 34, 35, 36, 37]
    sheet = spreadsheet.worksheet('Research Process')
    # for row in range(3, 38):
        # if row not in skip_list:
    for row in rows:
        prompt = sheet.acell('B%d' % row).value
        not_agree_dict = {'C': 'N/A or DID NOT DO THIS', 'D': 'STRONGLY DISAGREE', 'E': 'DISAGREE', 'F': 'NEITHER AGREE NOR DISAGREE'}
        for column in not_agree_dict.keys():
            val = sheet.acell('%c%d' % (column,row)).value
            if val: # if student answered anything other than agree or strongly agree, print the response and the propmt
                print not_agree_dict[column] + ": " + prompt
                not_agree_count += 1
    print not_agree_count

def FreeResponses():
    DTRFreeResponses()
    LIPFreeResponse()
    CollaborationFreeResponse()
    GrowthFreeResponse()

def Processes():
    DTRProcess()
    ResearchProcess()

def AnalyzeSelfAssessment():
    FreeResponses()
    Processes()

AnalyzeSelfAssessment()
