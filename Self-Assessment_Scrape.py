import gspread
from oauth2client.service_account import ServiceAccountCredentials

# authorize gspread
scope = ['https://spreadsheets.google.com/feeds']

credentials = ServiceAccountCredentials.from_json_keyfile_name('IDP_keys.json', scope)

gc = gspread.authorize(credentials)

# get prompts and responses to design, technology, and research sheets
def DTRFreeResponses(doc_id):
    prompt_list =[]
    response_list =[]
    spreadsheet = gc.open_by_key(doc_id) # if not found or won't open, go to spreadsheet (or folder it is contained in) and share  w/ client_email in .json keys file
    for page_number in range(3, 6):
        sheet = spreadsheet.get_worksheet(page_number)
        for row in range(3, 6):
            prompt = sheet.acell('A%d' % row).value
            response = sheet.acell('B%d' % row).value
            prompt_list.append(prompt)
            response_list.append(response)
    result = {'prompts': prompt_list, 'responses': response_list}
    return result

def LIPFreeResponse(doc_id):
    if doc_id != '1bxNRyLRmf0jvwnkom4hPOchLhA7MA_nrUFF6zrmUqZE' and doc_id != '1YfJN4Kk0hbAhGotqHMjECU2GNsh7NdPiP_U8NBPQAT0':
        prompt_list =[]
        response_list =[]
        spreadsheet = gc.open_by_key(doc_id)
        sheet = spreadsheet.worksheet('LIP')
        for row in range(5, 8):
            prompt = sheet.acell('A%d' % row).value
            response = sheet.acell('B%d' % row).value
            prompt_list.append(prompt)
            response_list.append(response)
        result = {'prompts': prompt_list, 'responses': response_list}
        return result
    return {'prompts': [], 'responses': []}


def CollaborationFreeResponse(doc_id):
    prompt_list =[]
    response_list =[]
    spreadsheet = gc.open_by_key(doc_id)
    sheet = spreadsheet.worksheet('Collaboration')
    for row in range(4, 7):
        prompt = sheet.acell('A%d' % row).value
        response = sheet.acell('B%d' % row).value
        prompt_list.append(prompt)
        response_list.append(response)
    result = {'prompts': prompt_list, 'responses': response_list}
    return result

def GrowthFreeResponse(doc_id):
    prompt_list =[]
    response_list =[]
    spreadsheet = gc.open_by_key(doc_id)
    sheet = spreadsheet.worksheet('Growth')
    for row in range(5, 8):
        prompt = sheet.acell('A%d' % row).value
        response = sheet.acell('B%d' % row).value
        prompt_list.append(prompt)
        response_list.append(response)
    result = {'prompts': prompt_list, 'responses': response_list}
    return result


def DTRProcess(doc_id):
    # not using 11, 12, 13, 16, 20, 32
    not_agree_count = 0
    rows = [2, 3, 4, 5, 6, 7, 8, 9, 10, 14, 15, 17, 18, 19,
    21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34]
    na_list = []
    strong_da_list = []
    da_list = []
    neither_list = []
    spreadsheet = gc.open_by_key(doc_id)
    sheet = spreadsheet.worksheet('DTR Process')
    for row in rows:
        prompt = sheet.acell('B%d' % row).value
        not_agree_dict = {'C': 'N/A or DID NOT DO THIS', 'D': 'STRONGLY DISAGREE', 'E': 'DISAGREE', 'F': 'NEITHER AGREE NOR DISAGREE'}
        for column in not_agree_dict.keys():
            val = sheet.acell('%c%d' % (column,row)).value
            if val: # if student answered anything other than agree or strongly agree, print the response and the propmt
                if column is 'C':
                    na_list.append(prompt)
                elif column is 'D':
                    strong_da_list.append(prompt)
                elif column is 'E':
                    da_list.append(prompt)
                else:
                    neither_list.append(prompt)
                # print not_agree_dict[column] + ": " + prompt
                not_agree_count += 1
    # print not_agree_count
    result = {'NA':na_list, 'Strongly Disagree':strong_da_list, 'Disagree':da_list, 'Neither Agree nor Disagree':neither_list}
    return result



def ResearchProcess(doc_id):
    # not using 2, 10, 11, 12, 16, 24, 25, 29
    not_agree_count = 0
    rows = [3, 4, 5, 6, 7, 8, 9, 13, 14, 15, 17, 18, 19, 20, 21,
    22, 23, 26, 27, 28, 30, 31, 32, 33, 34, 35, 36, 37]
    na_list = []
    strong_da_list = []
    da_list = []
    neither_list = []
    spreadsheet = gc.open_by_key(doc_id)
    sheet = spreadsheet.worksheet('Research Process')
    for row in rows:
        prompt = sheet.acell('B%d' % row).value
        not_agree_dict = {'C': 'N/A or DID NOT DO THIS', 'D': 'STRONGLY DISAGREE', 'E': 'DISAGREE', 'F': 'NEITHER AGREE NOR DISAGREE'}
        for column in not_agree_dict.keys():
            val = sheet.acell('%c%d' % (column,row)).value
            if val: # if student answered anything other than agree or strongly agree, print the response and the propmt
                if column is 'C':
                    na_list.append(prompt)
                elif column is 'D':
                    strong_da_list.append(prompt)
                elif column is 'E':
                    da_list.append(prompt)
                else:
                    neither_list.append(prompt)
                # print not_agree_dict[column] + ": " + prompt
                not_agree_count += 1
    # print not_agree_count
    result = {'NA':na_list, 'Strongly Disagree':strong_da_list, 'Disagree':da_list, 'Neither Agree nor Disagree':neither_list}
    return result

# output worksheet
output_spreadsheet = gc.open_by_key('1qTcgo015pJfn-w-r4SyvrOYmQy3HC_XzzQy1TISsgqk')

# Collect free response data and import into result spreadsheet
def FreeResponses(doc_id, name):
    dtr_dict = DTRFreeResponses(doc_id)
    lip_dict = LIPFreeResponse(doc_id)
    collab_dict = CollaborationFreeResponse(doc_id)
    growth_dict = GrowthFreeResponse(doc_id)

    free_response_prompts = dtr_dict['prompts'] + lip_dict['prompts'] + collab_dict['prompts'] + growth_dict['prompts']
    free_response_responses = dtr_dict['responses'] + lip_dict['responses'] + collab_dict['responses'] + growth_dict['responses']

    wks = output_spreadsheet.add_worksheet(title = name,  rows = "100", cols="20")
    if doc_id != '1bxNRyLRmf0jvwnkom4hPOchLhA7MA_nrUFF6zrmUqZE' and doc_id != '1YfJN4Kk0hbAhGotqHMjECU2GNsh7NdPiP_U8NBPQAT0': # Exceptions for Sehmon and Slim - different format
        cell_list_a = wks.range('A1:A18')
        cell_list_b = wks.range('B1:B18')
        start = 1
        end = 18
    else:
        print len(free_response_prompts)
        print free_response_prompts
        cell_list_a = wks.range('A1:A15')
        cell_list_b = wks.range('B1:B15')
        start = 1
        end = 15

    cell_list_a[0].value = 'Prompts'
    cell_list_b[0].value = 'Your Response'


    for cell in range(start, end):
        cell_list_a[cell].value = free_response_prompts[cell]
        cell_list_b[cell].value = free_response_responses[cell]

    wks.update_cells(cell_list_a)
    wks.update_cells(cell_list_b)



def Processes(doc_id, name):
    dtr_dict = DTRProcess(doc_id)
    research_dict = ResearchProcess(doc_id)
    agreement_dict = ['Strongly Disagree', 'Disagree', 'Neither Agree nor Disagree', 'NA']
    prompt_list = [agreement_dict[0]] + dtr_dict[agreement_dict[0]] + research_dict[agreement_dict[0]] + ['', agreement_dict[1]] + dtr_dict[agreement_dict[1]] + research_dict[agreement_dict[1]] + ['', agreement_dict[2]] + dtr_dict[agreement_dict[2]] + research_dict[agreement_dict[2]] + ['', agreement_dict[3]] + dtr_dict[agreement_dict[3]] + research_dict[agreement_dict[3]]
    start = 0
    end = len(prompt_list)
    wks = output_spreadsheet.worksheet(name)
    cell_range = 'E2:E' + str(end+1)
    cell_list = wks.range(cell_range)
    count = 0
    wks.update_acell('F1', 'Choose at most 5 behaviors to prioritize this quarter')
    wks.update_acell('G1', 'Actionable Step to Improve')
    wks.update_acell('H1', 'What made this difficult last quarter?')
    for cell in cell_list:
        cell.value = prompt_list[count]
        count += 1
    wks.update_cells(cell_list)


name_key_dict = {
'Nneoma': '1a4-t-LI_IwfcM02byKARwBcU5-nIEMzFy_ohy_rtp1Q',
'Leesha': '1EZPz5BdxfYThDkiB599awrZvI_TVdYJiTjgA8Dk48qY',
'Allison Lu': '17nLBftemvRWuGk6JbfCoASlQeK99kyfdhKON8br_5eE',
'Megan': '1LroUxoSOPZDa_1H_MBKzAUOVvbjm7NFGri-kJGLvC_Q',
'Allisun': '1yZa7YB0Q1A51HNfdR6JuBqH_-TehEZ9no44MuhmUrcg',
'Jennie': '1gg3fhYYjuyVWkNHF6gT3Rh4nkeIlnGFiPmq8G9Z-XvA',
'Meggy G': '1bAhVEe_E8IGJJzefQC12Bb6jX2KY4fdVeoCFgouxRMo',
'Olivia': '1zkWZT9F0lBvGH2TYlyHauE2JXr7Iiywe_ypdhSqZmyY',
'Ryan Louie': '1r71iz1MBOLSB03t5AZA5gyLQhP9okHEwmH99sNVq660',
'Armaan': '1C8oID5OYgQ7SNR_emeP5DCmrsrdGd6lNkvO_yxsXbqg',
'Josh': '1mAkhLb3F7pYiqZ9fkawBemTQjKNoCYYmiDotvI0dzF0',
'Slim': '1YfJN4Kk0hbAhGotqHMjECU2GNsh7NdPiP_U8NBPQAT0',
'Garrett': '1G5ozgM_EPHBsKcsLgDJuH3lEpw6SjNXX-QSxis0h3nw',
'yk': '14RWC5UKF_YkX45V0GAtsmp7P0UpK4lhNGSCI2Jgy4A0',
'Grace': '1ZvSzlztZ5-8YgcsXHWX9GT1JpqzjJKSD56WT_6uG8ys',
'Sehmon': '1bxNRyLRmf0jvwnkom4hPOchLhA7MA_nrUFF6zrmUqZE',
}


# Output for free responses
# for name in name_key_dict:
#     print name
#     FreeResponses(name_key_dict[name], name)


# Output for structured responses
# for name in name_key_dict:
#     Processes(name_key_dict[name], name)
#
def PreAndPostSurvey():
    pre_post_spreadsheet = gc.open_by_key('1GxUBwC7GJU9DXpj4387JGz6cIJLHJ5RYStF8Fz06uXA')

    name_array = [
    'Leesha', 'Allison Lu', 'Megan',
    'Meggy G', 'Olivia', 'Ryan Louie',
    'Armaan', 'Josh', 'Slim', 'Garrett', 'yk',
    'Grace', 'Sehmon', 'Allisun', 'Jennie', 'Nneoma'
    ]

    # make a new worksheet for each returning student
    for name in name_array:
        wks = pre_post_spreadsheet.add_worksheet(title = name,  rows = "60", cols="20")

        # pre survey questions
        wks.update_acell('A1', 'Pre-Activity Questions')
        wks.update_acell('A2', 'What were some of your weaknesses last quarter that you want to focus on improving this winter?')
        wks.update_acell('A3', 'How confident do you feel about your ability to grow in those areas?')

        # post survey questions
        wks.update_acell('D1', 'Post-Activity Questions')
        wks.update_acell('D2', 'After refelcting on your self-assessment, have your priorities for this quarter changed?')
        wks.update_acell('D3', 'What are some actionable steps you can take to work on the weaknesses you listed in cell B3 this quarter?')
        wks.update_acell('D4', 'How confident do you feel about your ability to work on your areas of weakness?')



def Processes():
    DTRProcess()
    ResearchProcess()


def AnalyzeSelfAssessment():
    FreeResponses()
    Processes()

############ RUN 'EM ##########

# for name in name_key_dict:
#     print name
#     FreeResponses(name_key_dict[name], name)

PreAndPostSurvey()
