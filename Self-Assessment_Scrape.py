import gspread
from oauth2client.service_account import ServiceAccountCredentials

# authorize gspread
scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('IDP_keys.json', scope)

gc = gspread.authorize(credentials)


# open self-assessment worksheet
spreadsheet = gc.open_by_key('1a4-t-LI_IwfcM02byKARwBcU5-nIEMzFy_ohy_rtp1Q')
# won't open...
    # resolution!: go to spreadsheet and share sheet w/ client_email in .json keys file

# get prompts and responses to design, technology, and research sheets
def DTRFreeResponses():
    prompt_list =[]
    response_list =[]
    for page_number in range(3, 6):
        sheet = spreadsheet.get_worksheet(page_number)
        for row in range(3, 6):
            prompt = sheet.acell('A%d' % row).value
            response = sheet.acell('B%d' % row).value
            # print prompt + ': ' + response
            prompt_list.append(prompt)
            response_list.append(response)
    result = {'prompts': prompt_list, 'responses': response_list}
    return result

def LIPFreeResponse():
    prompt_list =[]
    response_list =[]
    sheet = spreadsheet.worksheet('LIP')
    for row in range(5, 8):
        prompt = sheet.acell('A%d' % row).value
        response = sheet.acell('B%d' % row).value
        # print prompt + ': ' + response
        prompt_list.append(prompt)
        response_list.append(response)
    result = {'prompts': prompt_list, 'responses': response_list}
    return result


def CollaborationFreeResponse():
    prompt_list =[]
    response_list =[]
    sheet = spreadsheet.worksheet('Collaboration')
    for row in range(4, 7):
        prompt = sheet.acell('A%d' % row).value
        response = sheet.acell('B%d' % row).value
        # print prompt + ': ' + response
        prompt_list.append(prompt)
        response_list.append(response)
    result = {'prompts': prompt_list, 'responses': response_list}
    return result

def GrowthFreeResponse():
    prompt_list =[]
    response_list =[]
    sheet = spreadsheet.worksheet('Growth')
    for row in range(5, 8):
        prompt = sheet.acell('A%d' % row).value
        response = sheet.acell('B%d' % row).value
        # print prompt + ': ' + response
        prompt_list.append(prompt)
        response_list.append(response)
    result = {'prompts': prompt_list, 'responses': response_list}
    return result


def DTRProcess():
    # would be usefull to tally which cells are called on often b/w assessments
    # not using 11, 12, 13, 16, 20, 32
    not_agree_count = 0
    rows = [2, 3, 4, 5, 6, 7, 8, 9, 10, 14, 15, 17, 18, 19,
    21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34]
    sheet = spreadsheet.worksheet('DTR Process')
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
    rows = [3, 4, 5, 6, 7, 8, 9, 13, 14, 15, 17, 18, 19, 20, 21,
    22, 23, 26, 27, 28, 30, 31, 32, 33, 34, 35, 36, 37]
    sheet = spreadsheet.worksheet('Research Process')
    for row in rows:
        prompt = sheet.acell('B%d' % row).value
        not_agree_dict = {'C': 'N/A or DID NOT DO THIS', 'D': 'STRONGLY DISAGREE', 'E': 'DISAGREE', 'F': 'NEITHER AGREE NOR DISAGREE'}
        for column in not_agree_dict.keys():
            val = sheet.acell('%c%d' % (column,row)).value
            if val: # if student answered anything other than agree or strongly agree, print the response and the propmt
                print not_agree_dict[column] + ": " + prompt
                not_agree_count += 1
    print not_agree_count

# output worksheet
output_spreadsheet = gc.open_by_key('1qTcgo015pJfn-w-r4SyvrOYmQy3HC_XzzQy1TISsgqk')

def FreeResponses():
    dtr_dict = DTRFreeResponses()
    lip_dict = LIPFreeResponse()
    collab_dict = CollaborationFreeResponse()
    growth_dict = GrowthFreeResponse()

    free_response_prompts = dtr_dict['prompts'] + lip_dict['prompts'] + collab_dict['prompts'] + growth_dict['prompts']
    free_response_responses = dtr_dict['responses'] + lip_dict['responses'] + collab_dict['responses'] + growth_dict['responses']

    wks = output_spreadsheet.get_worksheet(0)
    cell_list_a = wks.range('A2:A18')
    cell_list_b = wks.range('B2:B18')

    for cell in range(0, 17):
        cell_list_a[cell].value = free_response_prompts[cell]
        cell_list_b[cell].value = free_response_responses[cell]
    wks.update_cells(cell_list_a)
    wks.update_cells(cell_list_b)

FreeResponses()

def Processes():
    DTRProcess()
    ResearchProcess()

def AnalyzeSelfAssessment():
    FreeResponses()
    Processes()



# AnalyzeSelfAssessment()


# putting the data into a spreadsheet

# # # Update cell range
# wks = output_spreadsheet.get_worksheet(0)
# cell_list = wks.range('A1:A17')
# for cell in cell_list:
#     cell.value = 'dfjkasd'
#
# wks.update_cells(cell_list) # Update in batch



# next steps:
    # put the output in a format that'll be easy for people to read
    # collect data on highest instances of "not agreeing"
    # run for each self-assessment
    # make faster if time
