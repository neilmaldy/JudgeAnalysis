import json
import csv
import pprint
import xlsxwriter
import argparse
from os import chdir, path
from collections import defaultdict, Counter
from time import sleep
from sys import exit

judge_id_to_name = defaultdict(str)

def max_column_width(x: int, y: str) -> int:
    return max(x, len(str(y)))


def append_row_2(worksheet, list_to_append, cell_format):
    global judge_id_to_name
    if len(judge_id_to_name):
        temp_list = list_to_append.copy()
        for i, item in enumerate(temp_list):
            try:
                judge_id = item.split()[0]
                if '-' in judge_id:
                    temp_text = item + ' ' + judge_id_to_name[judge_id]
                    list_to_append[i] = temp_text
            except Exception as e:
                pass
    try:
        worksheet.write_row(worksheet.row_counter, 0, list_to_append, cell_format)
        if len(list_to_append) > len(worksheet.column_widths):
            worksheet.column_widths.extend([1] * (len(list_to_append) - len(worksheet.column_widths)))
        if len(worksheet.column_widths) > len(list_to_append):
            list_to_append.extend([''] * (len(worksheet.column_widths) - len(list_to_append)))
        if list_to_append:
            worksheet.column_widths = list(map(max_column_width, worksheet.column_widths, list_to_append))
    except AttributeError:
        worksheet.row_counter = 0
        worksheet.column_widths = [len(x) for x in list_to_append]
        worksheet.write_row(worksheet.row_counter, 0, list_to_append, cell_format)
    worksheet.row_counter += 1
    return worksheet.row_counter


def set_column_widths(worksheet):
    if hasattr(worksheet, 'column_widths'):
        last_column_id = None
        last_column_width = None
        for column_id, column_width in enumerate(worksheet.column_widths):
            worksheet.set_column(column_id, column_id, min(20, column_width*0.85)+1.0)
            last_column_id = column_id
            last_column_width = column_width
        if False and last_column_id and last_column_width:
            worksheet.set_column(last_column_id, last_column_id, last_column_width)
    return

def main():

    parser = argparse.ArgumentParser(description='analyze_scores.py')

    debugit = False
    wb = xlsxwriter.Workbook()
    speed_sheet = wb.add_worksheet('Speed')
    miss_sheet = wb.add_worksheet('Misses')
    break_sheet = wb.add_worksheet('Breaks')
    presentation_sheet = wb.add_worksheet('Presentation')
    difficulty_sheet = wb.add_worksheet('Difficulty')

    data_cell_format = wb.add_format({'border': 1})
    bold_cell_format = wb.add_format({'bold': True})
    parser.add_argument('filename', metavar='filename', type=str, nargs='?', default='', help='Scoring file name')
    parser.add_argument('-a', '--anonymous', help='Do not include entry numbers', action='store_true')
    try:
        args = parser.parse_args()
        if args.filename:
            print("args.filename: ", args.filename)
            sleep(0.2)
            dirname = path.dirname(args.filename)
            if dirname:
                chdir(dirname)
            filename = path.basename(args.filename)
            print("Filename: ", filename)
        else:
            # filename = 'CompetitionScores_US Trials 2025_2025-02-19_22-35-05.tsv'
            # args.anonymous = True
            # filename = 'ZCompetitionScores_Zero Hour 2025_2025-01-18_20-04-06.tsv'
            # filename = 'CompetitionScores_YMCA Super Skipper Judge Training_2025-02-08_01-51-28.tsv'
            # filename = 'FCompetitionScores_Fast Feet and Freestyle Faceoff_2025-01-18_20-04-25.tsv'
            print('No scoring filename provided')
            input('press enter to quit')
            exit()
    except Exception as e:
        print(str(e))
        print("Problem with scoring file")
        input('press enter to quit')
        exit()
    print('Reading file: ', filename)
    sleep(0.2)
    # input('press enter')
    try:
        file = open(filename, 'r')
        dict_reader = csv.DictReader(file, delimiter='\t')
        data = [row for row in dict_reader]
        file.close()
        print('File read')
        sleep(0.2)
        # input('press enter')
    except Exception as e:
        print(str(e))
        print("Problem reading scoring file")
        input('press enter to quit')
        exit()

    entry_to_teamname = defaultdict(str)
    if path.exists('entries.csv'):
        try:
            print("Reading entries.csv")
            sleep(0.2)
            # input('press enter')
            file = open('entries.csv', 'r')
            dict_reader = csv.DictReader(file)
            for row in dict_reader:
                entry_to_teamname[row['EntryNumber']] = row['TeamName'] + ' r:' + str(row['Rank'])
            file.close()
            print("entries.csv read")
            sleep(0.2)
            # input('press enter')
        except Exception as e:
            print(str(e))
            print("Problem reading entries.csv")
            sleep(0.2)
    else:
        print("entries.csv not found, team names will not be included")

    if path.exists('judges.tsv'):
        try:
            print("Reading judges.tsv")
            sleep(0.2)
            # input('press enter')
            file = open('judges.tsv', 'r')
            dict_reader = csv.DictReader(file, delimiter='\t')
            for row in dict_reader:
                judge_id_to_name[row['JudgeID']] = row['JudgeName']
            file.close()
            print("judges.tsv read")
            sleep(0.2)
            # input('press enter')
        except Exception as e:
            print(str(e))
            print("Problem reading judges.tsv")
            sleep(0.2)
    else:
        print("judges.tsv not found, judge names will not be included")

    scores = {}
    adjustments = {}
    missing_station_ids = set()
    skipped_events = set()
    for row in data:
        try:
            judgedata = row['JudgeScoreDataString']
            competition_name = row['CompetitionName']
            session_name = row['SessionName']
            session_id = row['SessionID']
            entry_number = row['EntryNumber']
            event_definition_abbr = row['EventDefinitionAbbr']
            if event_definition_abbr in ['DDCF', 'SCTF']:
                if event_definition_abbr not in skipped_events:
                    print("Skipping event: ", event_definition_abbr)
                    skipped_events.add(event_definition_abbr)
                continue
            station_id = str(row['StationID'])
            station_sequence = str(row['StationSequence'])
            score_sequence = str(row['ScoreSequence'])
            if not station_id:
                station_id = '0000'
                if entry_number not in missing_station_ids:
                    missing_station_ids.add(entry_number)
                    print("Station ID not found for entry number: ", entry_number)
            judge_id = station_id + '-' + score_sequence
            is_scored = row['IsScored']
            total_score = row['TotalScore']
            entry_is_scored = row['EntryIsScored']
            is_locked = row['IsLocked']
            judge_is_scored = row['JudgeIsScored']
            if is_scored == 'True':
                judge_score_data = json.loads(row['JudgeScoreDataString'])
                if 'JudgeResults' not in judge_score_data:
                    print("No judge results for entry number: " + entry_number + " judge_id: " + judge_id)
                    continue
                #todo check for DDCF and TeamShow
                judge_meta_data = judge_score_data['JudgeResults']['meta']
                judge_tally_data = judge_score_data['TallySheet']['tally']
                if judge_meta_data['judgeTypeId'] == 'P' and 'MarkSheet' in judge_score_data:
                    adjustments[(entry_number, judge_id)] = []
                    for mark in judge_score_data['MarkSheet']['marks']:
                        if 'Adj' in mark['schema']:
                            adjustments[(entry_number, judge_id)].append((mark['schema']))
                judge_results = judge_score_data['JudgeResults']['result']
                if judge_meta_data['judgeTypeId'] not in scores:
                    scores[judge_meta_data['judgeTypeId']] = {}
                if entry_number not in scores[judge_meta_data['judgeTypeId']]:
                    scores[judge_meta_data['judgeTypeId']][entry_number] = []
                scores[judge_meta_data['judgeTypeId']][entry_number].append((event_definition_abbr, judge_id, judge_tally_data, judge_results))
                # pprint.pprint(judge_tally_data)
                # pprint.pprint(judge_results)
        except Exception as e:
            print(str(e))
            print("Problem with entry number: ", entry_number)

    print("Scores parsed")
    sleep(0.2)

    for judge_type in ['Dr', 'Dm', 'Dp', 'Db', 'Da', 'Dj', 'Dt', 'P', 'T', 'Shj', 'S']:
        if judge_type not in scores:
            if debugit: print("No data for judge type: ", judge_type)
            scores[judge_type] = {}

    if debugit:
        for judge_type_id in scores:
            for entry_number in scores[judge_type_id]:
                if debugit:
                    print(judge_type_id, entry_number)
                    for event_definition_abbr, judge_id, judge_tally_data, judge_results in scores[judge_type_id][entry_number]:
                        print(event_definition_abbr, judge_id, judge_tally_data)
                        # pprint.pprint(judge_tally_data)
                        pprint.pprint(judge_results)

    misses_station_entry_rows = {}
    presentation_station_entry_rows = {}
    breaks_station_entry_rows = {}
    speed_station_entry_rows = {}

    for speed_judge_type in ['Shj', 'S']:
        for entry_number in scores[speed_judge_type]:
            if debugit: print(entry_number)
            for event_definition_abbr, judge_id, judge_tally_data, judge_results in scores[speed_judge_type][entry_number]:
                station_id = judge_id.split('-')[0]
                if station_id == '0000': continue
                if station_id not in speed_station_entry_rows:
                    speed_station_entry_rows[station_id] = {}
                    speed_station_entry_rows[station_id]['judge_ids'] = []
                    speed_station_entry_rows[station_id]['entries'] = {}
                    speed_station_entry_rows[station_id]['entry_types'] = {}
                if judge_id not in speed_station_entry_rows[station_id]['judge_ids']:
                    speed_station_entry_rows[station_id]['judge_ids'].append(judge_id)
                if entry_number not in speed_station_entry_rows[station_id]['entries']:
                    speed_station_entry_rows[station_id]['entries'][entry_number] = {}
                if entry_number not in speed_station_entry_rows[station_id]['entry_types']:
                    speed_station_entry_rows[station_id]['entry_types'][entry_number] = event_definition_abbr
                speed_station_entry_rows[station_id]['entries'][entry_number][judge_id] = judge_tally_data['step']
    print("Speed data parsed")

    for entry_number in scores['P']:
        if debugit: print(entry_number)
        for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(scores['P'][entry_number], key=lambda x: x[1]):
            station_id = judge_id.split('-')[0]
            if station_id == '0000': continue
            if station_id not in misses_station_entry_rows:
                misses_station_entry_rows[station_id] = {}
                misses_station_entry_rows[station_id]['judge_ids'] = []
                misses_station_entry_rows[station_id]['judge_types'] = {}
                misses_station_entry_rows[station_id]['entries'] = {}
                misses_station_entry_rows[station_id]['entry_types'] = {}
            if judge_id not in misses_station_entry_rows[station_id]['judge_ids']:
                misses_station_entry_rows[station_id]['judge_ids'].append(judge_id)
            if judge_id not in misses_station_entry_rows[station_id]['judge_types']:
                misses_station_entry_rows[station_id]['judge_types'][judge_id] = 'P'
            if entry_number not in misses_station_entry_rows[station_id]['entries']:
                misses_station_entry_rows[station_id]['entries'][entry_number] = {}
            if entry_number not in misses_station_entry_rows[station_id]['entry_types']:
                misses_station_entry_rows[station_id]['entry_types'][entry_number] = event_definition_abbr
            misses_station_entry_rows[station_id]['entries'][entry_number][judge_id] = judge_results['nm']

            if station_id not in presentation_station_entry_rows:
                presentation_station_entry_rows[station_id] = {}
                presentation_station_entry_rows[station_id]['judge_ids'] = []
                presentation_station_entry_rows[station_id]['judge_types'] = {}
                presentation_station_entry_rows[station_id]['entries'] = {}
                presentation_station_entry_rows[station_id]['entry_types'] = {}
                presentation_station_entry_rows[station_id]['judge_stats'] = {}
            if judge_id not in presentation_station_entry_rows[station_id]['judge_ids']:
                presentation_station_entry_rows[station_id]['judge_ids'].append(judge_id)
            if judge_id not in presentation_station_entry_rows[station_id]['judge_types']:
                presentation_station_entry_rows[station_id]['judge_types'][judge_id] = 'P'
            if entry_number not in presentation_station_entry_rows[station_id]['entries']:
                presentation_station_entry_rows[station_id]['entries'][entry_number] = {}
                presentation_station_entry_rows[station_id]['entries'][entry_number]['p_list'] = []
            if entry_number not in presentation_station_entry_rows[station_id]['entry_types']:
                presentation_station_entry_rows[station_id]['entry_types'][entry_number] = event_definition_abbr
            adjustment_counts = Counter(adjustments.get((entry_number, judge_id), []))
            e_adjustments = adjustment_counts.get('entPlusAdj', 0) - adjustment_counts.get('entMinusAdj', 0)
            f_adjustments = adjustment_counts.get('formPlusAdj', 0) - adjustment_counts.get('formMinusAdj', 0)
            m_adjustments = adjustment_counts.get('musicPlusAdj', 0) - adjustment_counts.get('musicMinusAdj', 0)
            c_adjustments = adjustment_counts.get('creaPlusAdj', 0) - adjustment_counts.get('creaMinusAdj', 0)
            v_adjustments = adjustment_counts.get('variPlusAdj', 0) - adjustment_counts.get('variMinusAdj', 0)
            total_adjustments = e_adjustments + f_adjustments + m_adjustments + c_adjustments + v_adjustments
            presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id] = (round(judge_results['p'], 2), judge_tally_data['ent'], judge_tally_data['form'], judge_tally_data['music'], judge_tally_data['crea'], judge_tally_data['vari'], e_adjustments, f_adjustments, m_adjustments, c_adjustments, v_adjustments, total_adjustments)
            presentation_station_entry_rows[station_id]['entries'][entry_number]['p_list'].append(round(judge_results['p'], 2))
            if judge_id not in presentation_station_entry_rows[station_id]['judge_stats']:
                presentation_station_entry_rows[station_id]['judge_stats'][judge_id] = []
            presentation_station_entry_rows[station_id]['judge_stats'][judge_id].append((round(judge_results['p'], 2), judge_tally_data['ent'], judge_tally_data['form'], judge_tally_data['music'], judge_tally_data['crea'], judge_tally_data['vari']))
    for entry_number in scores['T']:
        if debugit: print(entry_number)
        for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(scores['T'][entry_number], key=lambda x: x[1]):
            station_id = judge_id.split('-')[0]
            if station_id == '0000': continue
            if station_id not in misses_station_entry_rows:
                misses_station_entry_rows[station_id] = {}
                misses_station_entry_rows[station_id]['judge_ids'] = []
                misses_station_entry_rows[station_id]['judge_types'] = {}
                misses_station_entry_rows[station_id]['entries'] = {}
                misses_station_entry_rows[station_id]['entry_types'] = {}
            if judge_id not in misses_station_entry_rows[station_id]['judge_ids']:
                misses_station_entry_rows[station_id]['judge_ids'].append(judge_id)
            if judge_id not in misses_station_entry_rows[station_id]['judge_types']:
                misses_station_entry_rows[station_id]['judge_types'][judge_id] = 'T'
            if entry_number not in misses_station_entry_rows[station_id]['entries']:
                misses_station_entry_rows[station_id]['entries'][entry_number] = {}
            if entry_number not in misses_station_entry_rows[station_id]['entry_types']:
                misses_station_entry_rows[station_id]['entry_types'][entry_number] = event_definition_abbr
            misses_station_entry_rows[station_id]['entries'][entry_number][judge_id] = judge_results['nm']

            if event_definition_abbr in ['SRIF', 'SRPF', 'SRTF', 'WHPF']:
                if station_id not in breaks_station_entry_rows:
                    breaks_station_entry_rows[station_id] = {}
                    breaks_station_entry_rows[station_id]['judge_ids'] = []
                    breaks_station_entry_rows[station_id]['judge_types'] = {}
                    breaks_station_entry_rows[station_id]['entries'] = {}
                    breaks_station_entry_rows[station_id]['entry_types'] = {}
                if judge_id not in breaks_station_entry_rows[station_id]['judge_ids']:
                    breaks_station_entry_rows[station_id]['judge_ids'].append(judge_id)
                if judge_id not in breaks_station_entry_rows[station_id]['judge_types']:
                    breaks_station_entry_rows[station_id]['judge_types'][judge_id] = 'T'
                if entry_number not in breaks_station_entry_rows[station_id]['entries']:
                    breaks_station_entry_rows[station_id]['entries'][entry_number] = {}
                if entry_number not in breaks_station_entry_rows[station_id]['entry_types']:
                    breaks_station_entry_rows[station_id]['entry_types'][entry_number] = event_definition_abbr
                breaks_station_entry_rows[station_id]['entries'][entry_number][judge_id] = judge_results['nb']

    for entry_number in scores['Dj']:
        if debugit: print(entry_number)
        for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(scores['Dj'][entry_number], key=lambda x: x[1]):
            station_id = judge_id.split('-')[0]
            if station_id == '0000': continue
            if event_definition_abbr in ['DDSF', 'DDPF']:
                if station_id not in breaks_station_entry_rows:
                    breaks_station_entry_rows[station_id] = {}
                    breaks_station_entry_rows[station_id]['judge_ids'] = []
                    breaks_station_entry_rows[station_id]['judge_types'] = {}
                    breaks_station_entry_rows[station_id]['entries'] = {}
                    breaks_station_entry_rows[station_id]['entry_types'] = {}
                if judge_id not in breaks_station_entry_rows[station_id]['judge_ids']:
                    breaks_station_entry_rows[station_id]['judge_ids'].append(judge_id)
                if judge_id not in breaks_station_entry_rows[station_id]['judge_types']:
                    breaks_station_entry_rows[station_id]['judge_types'][judge_id] = 'Dj'
                if entry_number not in breaks_station_entry_rows[station_id]['entries']:
                    breaks_station_entry_rows[station_id]['entries'][entry_number] = {}
                if entry_number not in breaks_station_entry_rows[station_id]['entry_types']:
                    breaks_station_entry_rows[station_id]['entry_types'][entry_number] = event_definition_abbr
                breaks_station_entry_rows[station_id]['entries'][entry_number][judge_id] = judge_tally_data['break']

    sr_scores_station_entry_rows = {}
    dd_scores_station_entry_rows = {}
    for judge_type_id in ['Dr', 'Dm', 'Dp', 'Db', 'Da', 'Dj', 'Dt']:
        for entry_number in scores[judge_type_id]: # Dr Dm Dp Db Da Dj Dt
            if debugit: print(entry_number)
            for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(scores[judge_type_id][entry_number], key=lambda x: x[1]):
                station_id = judge_id.split('-')[0]
                if station_id == '0000': continue
                if event_definition_abbr in ['SRIF', 'SRPF', 'SRTF', 'WHPF'] and judge_type_id in ['Dr', 'Dm', 'Dp', 'Db', 'Da']:
                    if station_id not in sr_scores_station_entry_rows:
                        sr_scores_station_entry_rows[station_id] = {}
                        sr_scores_station_entry_rows[station_id]['judge_type'] = {}
                    if judge_type_id not in sr_scores_station_entry_rows[station_id]['judge_type']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id] = {}
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'] = []
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'] = {}
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'] = {}
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'] = {}
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'] = {}
                    if judge_id not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'].append(judge_id)
                    if entry_number not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][entry_number] = []
                    sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][entry_number].append(round(judge_results['d'],2))
                    if entry_number not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number] = {}
                    if entry_number not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][entry_number] = event_definition_abbr
                    temp_dict = judge_tally_data
                    temp_dict.pop('rep', None)
                    temp_dict.pop('break', None)
                    temp_dict['d'] = round(judge_results['d'],2)
                    if 'columns' not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'] = sorted(temp_dict.keys())
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'].append('Avg Clicks/Heat')
                    temp_list = []
                    total_score = 0
                    for key in sorted(temp_dict.keys()):
                        temp_list.append(temp_dict[key])
                        if 'diff' in key:
                            total_score += temp_dict[key]
                    temp_dict['Total'] = total_score
                    temp_list.append(total_score)
                    sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number][judge_id] = tuple(temp_list)
                    if judge_id not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id] = {}
                        for key in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns']:
                            sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][key] = 0
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]['heat_count'] = 0  
                    if total_score > 0:
                        for key in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]:
                            if key == 'heat_count':
                                sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][key] += 1
                            elif key == 'Avg Clicks/Heat':
                                sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][key] += temp_dict['Total']
                            else:
                                sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][key] += temp_dict[key]

                if event_definition_abbr in ['DDSF', 'DDPF'] and judge_type_id in ['Dj', 'Dt']:
                    if station_id not in dd_scores_station_entry_rows:
                        dd_scores_station_entry_rows[station_id] = {}
                        dd_scores_station_entry_rows[station_id]['judge_type'] = {}
                    if judge_type_id not in dd_scores_station_entry_rows[station_id]['judge_type']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id] = {}
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'] = []
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'] = {}
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'] = {}
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'] = {}
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'] = {}
                    if judge_id not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'].append(judge_id)
                    if entry_number not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][entry_number] = []
                    dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][entry_number].append(round(judge_results['d'],2))
                    if entry_number not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number] = {}
                    if entry_number not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][entry_number] = event_definition_abbr
                    temp_dict = {'d': round(judge_results['d'],2)}
                    temp_dict.update(judge_tally_data)
                    temp_dict.pop('rep', None)
                    temp_dict.pop('break', None)
                    temp_dict['d'] = round(judge_results['d'],2)
                    if 'columns' not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'] = list(temp_dict.keys())
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'].append('Avg Clicks/Heat')
                    temp_list = []
                    total_score = 0
                    for key in sorted(temp_dict.keys()):
                        temp_list.append(temp_dict[key])
                        if 'diff' in key:
                            total_score += temp_dict[key]
                    temp_dict['Total'] = total_score
                    temp_list.append(total_score)
                    dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number][judge_id] = tuple(temp_list)
                    if judge_id not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id] = {}
                        for key in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns']:
                            dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][key] = 0
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]['heat_count'] = 0  
                    if total_score > 0:
                        for key in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]:
                            if key == 'heat_count':
                                dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][key] += 1
                            elif key == 'Avg Clicks/Heat':
                                dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][key] += temp_dict['Total']
                            else:
                                dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][key] += temp_dict[key]


    with open('output.csv', 'w') as f:
        if debugit: print("Speed\n")
        print("Speed\n", file=f)
        for station_id in speed_station_entry_rows:
            if station_id == '0000': continue
            if debugit: print('Station: ' + station_id)
            print('Station: ' + station_id + ' speed scores and cummulative difference from calculated score', file=f)
            append_row_2(speed_sheet, ['Station: ' + station_id + ' Speed scores and cummulative difference from calculated score'], bold_cell_format)

            cummulative_error = {}
            speed_station_entry_rows[station_id]['judge_ids'].sort()
            row = ['Entry Number']
            for judge_id in speed_station_entry_rows[station_id]['judge_ids']:
                row.append(judge_id)
                cummulative_error[judge_id] = 0
            row.append('Calculated Score')
            num_columns = len(row)
            for judge_id in speed_station_entry_rows[station_id]['judge_ids']:
                row.append(judge_id + ' Diff')
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            header_row = append_row_2(speed_sheet, row, bold_cell_format)

            for entry_number in speed_station_entry_rows[station_id]['entries']:
                if args.anonymous:
                    my_entry_number = 'Entry n'
                else:
                    my_entry_number = entry_number
                row = [my_entry_number + ' ' + speed_station_entry_rows[station_id]['entry_types'][entry_number] + ' ' + entry_to_teamname[entry_number]]
                speed_scores = []
                for judge_id in speed_station_entry_rows[station_id]['judge_ids']:
                    if judge_id in speed_station_entry_rows[station_id]['entries'][entry_number]:
                        row.append(speed_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                        speed_scores.append(speed_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                    else:
                        row.append('')
                        speed_scores.append(0)
                sorted_speed_scores = sorted(speed_scores)
                if len(sorted_speed_scores) == 3:
                    if sorted_speed_scores[1] - sorted_speed_scores[0] < sorted_speed_scores[2] - sorted_speed_scores[1]:
                        calculated_score = round((sorted_speed_scores[0] + sorted_speed_scores[1]) / 2.0, 1)
                    else:
                        calculated_score = round((sorted_speed_scores[1] + sorted_speed_scores[2]) / 2.0, 1)
                    row.append(calculated_score)
                    for speed_score in speed_scores:
                        row.append(abs(round(calculated_score - speed_score, 1)))
                elif len(sorted_speed_scores) > 3:
                    calculated_score = round(sum(sorted_speed_scores[1:-1]) / (len(sorted_speed_scores) - 2), 1)
                    row.append(calculated_score)
                    for speed_score in speed_scores:
                        row.append(abs(round(calculated_score - speed_score, 1)))
                else:
                    calculated_score = 0
                    row.append(calculated_score)
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                last_row = append_row_2(speed_sheet, row, data_cell_format)
                speed_sheet.conditional_format(last_row-1, 1, last_row-1, num_columns-1, {'type': '3_color_scale'})
            speed_sheet.conditional_format(header_row, num_columns, last_row-1, num_columns + 2, {'type': '2_color_scale', 'min_color': 'white', 'max_color': 'red'})
            if len(sorted_speed_scores) == 3:
                sum_columns = ['F', 'G', 'H']
            elif len(sorted_speed_scores) == 4:
                sum_columns = ['G', 'H', 'I', 'J']
            elif len(sorted_speed_scores) == 5:
                sum_columns = ['H', 'I', 'J', 'K', 'L']
            else:
                sum_columns = []
            if sum_columns:
                for column in sum_columns:
                    speed_sheet.write_formula(column + str(last_row + 1), '=SUM(' + column + str(header_row+1) + ':' + column + str(last_row) + ')')
                speed_sheet.conditional_format(last_row, num_columns, last_row, num_columns + len(sum_columns), {'type': '2_color_scale', 'min_color': 'white', 'max_color': 'red', 'min_value': 0})
            # speed_sheet.write_formula('F' + str(last_row + 1), '=SUM(F' + str(header_row+1) + ':F' + str(last_row) + ')')
            # speed_sheet.write_formula('G' + str(last_row + 1), '=SUM(G' + str(header_row+1) + ':G' + str(last_row) + ')')
            # speed_sheet.write_formula('H' + str(last_row + 1), '=SUM(H' + str(header_row+1) + ':H' + str(last_row) + ')')
            # speed_sheet.conditional_format(last_row, num_columns, last_row, num_columns + 2, {'type': '2_color_scale', 'min_color': 'white', 'max_color': 'red', 'min_value': 0})
            if debugit: print()
            print('', file=f)
            append_row_2(speed_sheet, [], data_cell_format)

        if debugit: print("Misses\n")
        print("Misses\n", file=f)
        for station_id in misses_station_entry_rows:
            if station_id == '0000': continue
            if debugit: print(station_id)
            print(station_id, file=f)
            current_row = append_row_2(miss_sheet, [station_id + ' Misses'], bold_cell_format)
            misses_station_entry_rows[station_id]['judge_ids'].sort()
            row = ['Entry Number']
            for judge_id in misses_station_entry_rows[station_id]['judge_ids']:
                row.append(judge_id + ' ' + misses_station_entry_rows[station_id]['judge_types'][judge_id])
            # row = 'Entry Number,' + ','.join(station_entry_rows[station_id]['judge_ids'])
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            header_row = append_row_2(miss_sheet, row, data_cell_format)
            num_columns = len(row)

            running_totals = {}

            # change row strings to lists
            for entry_number in misses_station_entry_rows[station_id]['entries']:
                if args.anonymous:
                    my_entry_number = 'Entry n'
                else:
                    my_entry_number = entry_number
                row = [my_entry_number + ' ' + misses_station_entry_rows[station_id]['entry_types'][entry_number] + ' ' + entry_to_teamname[entry_number]]
                for judge_id in misses_station_entry_rows[station_id]['judge_ids']:
                    if judge_id in misses_station_entry_rows[station_id]['entries'][entry_number]:
                        # row += ',' + str(misses_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                        row.append(misses_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                        if judge_id not in running_totals:
                            running_totals[judge_id] = []
                        running_totals[judge_id].append(misses_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                    else:
                        # row += ','
                        row.append('')
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                last_row = append_row_2(miss_sheet, row, data_cell_format)
            miss_sheet.conditional_format(header_row, 1, last_row-1, num_columns-1, {'type': '3_color_scale', 'min_color': '#80FF80', 'mid_color': '#FFFF80', 'max_color': '#FF8080'})
            row = ['Totals']
            for judge_id in misses_station_entry_rows[station_id]['judge_ids']:
                if judge_id in running_totals:
                    row.append(sum(running_totals[judge_id]))
                else:
                    row.append('')
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            last_row = append_row_2(miss_sheet, row, data_cell_format)
            miss_sheet.conditional_format(last_row-1, 1, last_row-1, num_columns-1, {'type': '3_color_scale', 'min_color': '#80FF80', 'mid_color': '#FFFF80', 'max_color': '#FF8080'})

            row = ['Averages']
            for judge_id in misses_station_entry_rows[station_id]['judge_ids']:
                if judge_id in running_totals:
                    row.append(round(sum(running_totals[judge_id])/len(running_totals[judge_id]), 2))
                else:
                    row.append('')
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            last_row = append_row_2(miss_sheet, row, data_cell_format)
            miss_sheet.conditional_format(last_row-1, 1, last_row-1, num_columns-1, {'type': '3_color_scale', 'min_color': '#80FF80', 'mid_color': '#FFFF80', 'max_color': '#FF8080'})

            if debugit: print()
            print('', file=f)
            append_row_2(miss_sheet, [], data_cell_format)

        if debugit: print("Breaks\n")
        print("Breaks\n", file=f)
        for station_id in breaks_station_entry_rows:
            if station_id == '0000': continue
            if debugit: print(station_id)
            print(station_id, file=f)
            append_row_2(break_sheet, [station_id + ' Breaks'], bold_cell_format)

            breaks_station_entry_rows[station_id]['judge_ids'].sort()
            row = ['Entry Number']
            for judge_id in breaks_station_entry_rows[station_id]['judge_ids']:
                row.append(judge_id + ' ' + breaks_station_entry_rows[station_id]['judge_types'][judge_id])
            # row = 'Entry Number,' + ','.join(station_entry_rows[station_id]['judge_ids'])
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            header_row = append_row_2(break_sheet, row, data_cell_format)
            num_columns = len(row)
            for entry_number in breaks_station_entry_rows[station_id]['entries']:
                if args.anonymous:
                    my_entry_number = 'Entry n'
                else:
                    my_entry_number = entry_number
                row = [my_entry_number + ' ' + breaks_station_entry_rows[station_id]['entry_types'][entry_number] + ' ' + entry_to_teamname[entry_number]]
                for judge_id in breaks_station_entry_rows[station_id]['judge_ids']:
                    if judge_id in breaks_station_entry_rows[station_id]['entries'][entry_number]:
                        row.append(breaks_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                        if judge_id not in running_totals:
                            running_totals[judge_id] = []
                        running_totals[judge_id].append(breaks_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                    else:
                        row.append('')
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                last_row = append_row_2(break_sheet, row, data_cell_format)
            break_sheet.conditional_format(header_row, 1, last_row-1, num_columns-1, {'type': '3_color_scale', 'min_color': '#80FF80', 'mid_color': '#FFFF80', 'max_color': '#FF8080'})
            row = ['Totals']
            for judge_id in breaks_station_entry_rows[station_id]['judge_ids']:
                if judge_id in running_totals:
                    row.append(sum(running_totals[judge_id]))
                else:
                    row.append('')
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            last_row = append_row_2(break_sheet, row, data_cell_format)
            break_sheet.conditional_format(last_row-1, 1, last_row-1, num_columns-1, {'type': '3_color_scale', 'min_color': '#80FF80', 'mid_color': '#FFFF80', 'max_color': '#FF8080'})

            row = ['Averages']
            for judge_id in breaks_station_entry_rows[station_id]['judge_ids']:
                if judge_id in running_totals:
                    row.append(round(sum(running_totals[judge_id])/len(running_totals[judge_id]), 2))
                else:
                    row.append('')
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            last_row = append_row_2(break_sheet, row, data_cell_format)
            break_sheet.conditional_format(last_row-1, 1, last_row-1, num_columns-1, {'type': '3_color_scale', 'min_color': '#80FF80', 'mid_color': '#FFFF80', 'max_color': '#FF8080'})

            if debugit: print()
            print('', file=f)
            append_row_2(break_sheet, [], data_cell_format)

        if debugit: print("Presentation\n")
        print("Presentation\n", file=f)

        running_totals = {}
        for station_id in misses_station_entry_rows:
            if station_id == '0000': continue
            if debugit: print(station_id)
            judge_scores = {}
            judge_sorted_scores = {}
            judge_scores_ranked = {}
            print(station_id, file=f)
            append_row_2(presentation_sheet, [station_id + ' Avg Presentation vs Judges Score'], bold_cell_format)
            presentation_station_entry_rows[station_id]['judge_ids'].sort()

            p_avg = {}
            for entry_number in presentation_station_entry_rows[station_id]['entries']:
                if len(presentation_station_entry_rows[station_id]['entries'][entry_number]['p_list']) > 0:
                    p_avg[entry_number] = round(sum(presentation_station_entry_rows[station_id]['entries'][entry_number]['p_list'])/len(presentation_station_entry_rows[station_id]['entries'][entry_number]['p_list']), 2)
                else:
                    p_avg[entry_number] = 0
            sorted_p_avg = sorted(p_avg.items(), key=lambda x: x[1], reverse=True)
            row = ['Entry', 'P avg']
            row.extend(presentation_station_entry_rows[station_id]['judge_ids'])
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            header_row = append_row_2(presentation_sheet, row, data_cell_format)
            num_columns = len(row)

            for entry_number, p_value in sorted_p_avg:
                if args.anonymous:
                    my_entry_number = 'Entry n'
                else:
                    my_entry_number = entry_number
                row = [my_entry_number + ' ' + presentation_station_entry_rows[station_id]['entry_types'][entry_number] + ' ' + entry_to_teamname[entry_number], p_avg[entry_number]]
                for judge_id in presentation_station_entry_rows[station_id]['judge_ids']:
                    if judge_id not in judge_scores:
                        judge_scores[judge_id] = {}
                        judge_scores_ranked[judge_id] = {}
                    if judge_id in presentation_station_entry_rows[station_id]['entries'][entry_number]:
                        row.append(presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][0])
                        judge_scores[judge_id][entry_number] = presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][0]
                    else:
                        row.append('')
                        judge_scores[judge_id][entry_number] = 0
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                last_row = append_row_2(presentation_sheet, row, data_cell_format)
            # presentation_sheet.conditional_format(header_row, 1, last_row-1, num_columns-1, {'type': '3_color_scale'})
            for column in range(1, num_columns):
                presentation_sheet.conditional_format(header_row, column, last_row-1, column, {'type': '3_color_scale'})
            if debugit: print()
            print('', file=f)
            append_row_2(presentation_sheet, [], data_cell_format)

            print(station_id + ' Relative Rankings', file=f)
            append_row_2(presentation_sheet, [station_id + ' Relative Rankings'], bold_cell_format)
            row = ['Entry Number', 'Rank']
            row.extend(presentation_station_entry_rows[station_id]['judge_ids'])
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            header_row = append_row_2(presentation_sheet, row, data_cell_format)
            num_columns = len(row)

            for judge_id in presentation_station_entry_rows[station_id]['judge_ids']:
                judge_sorted_scores[judge_id] = dict(sorted(judge_scores[judge_id].items(), key=lambda item: item[1], reverse=True))
                rank = 1
                for entry_number in judge_sorted_scores[judge_id]:
                    judge_scores_ranked[judge_id][entry_number] = rank
                    rank += 1
            rank = 1
            for entry_number, p in sorted_p_avg:
                if args.anonymous:
                    my_entry_number = 'Entry n'
                else:
                    my_entry_number = entry_number
                row = [my_entry_number + ' ' + presentation_station_entry_rows[station_id]['entry_types'][entry_number] + ' ' + entry_to_teamname[entry_number], rank]
                for judge_id in presentation_station_entry_rows[station_id]['judge_ids']:
                    if judge_id in judge_scores_ranked:
                        row.append(judge_scores_ranked[judge_id][entry_number])
                    else:
                        row.append('')
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                last_row = append_row_2(presentation_sheet, row, data_cell_format)
                rank += 1
            presentation_sheet.conditional_format(header_row, 1, last_row-1, num_columns-1, {'type': 'data_bar'})
            if debugit: print()
            print('', file=f)
            append_row_2(presentation_sheet, [], data_cell_format)


            append_row_2(presentation_sheet, [station_id + ' Presentation Score Averages'], bold_cell_format)
            row = ['Judge', 'P avg', 'E avg', 'F avg', 'M avg', 'C avg', 'V avg']
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            header_row = append_row_2(presentation_sheet, row, data_cell_format)

            for judge_id in presentation_station_entry_rows[station_id]['judge_ids']:
                p_list = [x[0] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
                e_list = [x[1] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
                f_list = [x[2] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
                m_list = [x[3] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
                c_list = [x[4] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
                v_list = [x[5] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
                row = [judge_id]
                row.extend([round(sum(p_list)/len(p_list), 2), round(sum(e_list)/len(e_list), 2), round(sum(f_list)/len(f_list), 2), round(sum(m_list)/len(m_list), 2), round(sum(c_list)/len(c_list), 2), round(sum(v_list)/len(v_list), 2)])
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                last_row = append_row_2(presentation_sheet, row, data_cell_format)
            presentation_sheet.conditional_format(header_row, 1, last_row-1, 1, {'type': '3_color_scale'})
            presentation_sheet.conditional_format(header_row, 2, last_row-1, 2, {'type': '3_color_scale'})
            presentation_sheet.conditional_format(header_row, 3, last_row-1, 3, {'type': '3_color_scale'})
            presentation_sheet.conditional_format(header_row, 4, last_row-1, 4, {'type': '3_color_scale'})
            presentation_sheet.conditional_format(header_row, 5, last_row-1, 5, {'type': '3_color_scale'})
            presentation_sheet.conditional_format(header_row, 6, last_row-1, 6, {'type': '3_color_scale'})

            if debugit: print()
            print('', file=f)
            append_row_2(presentation_sheet, [], data_cell_format)

            append_row_2(presentation_sheet, [station_id + ' Presentation Score Details'], bold_cell_format)
            row = ['Entry Number', 'Judge', 'P', 'E', 'F', 'M', 'C', 'V', 'E adj', 'F adj', 'M adj', 'C adj', 'V adj', 'Total adj']
            if debugit: print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            header_row = append_row_2(presentation_sheet, row, data_cell_format)
            last_entry_row = header_row
            for entry_number in presentation_station_entry_rows[station_id]['entries']:
                for judge_id in presentation_station_entry_rows[station_id]['judge_ids']:
                    if args.anonymous:
                        my_entry_number = 'Entry n'
                    else:
                        my_entry_number = entry_number
                    row = [my_entry_number + ' ' + presentation_station_entry_rows[station_id]['entry_types'][entry_number] + ' ' + entry_to_teamname[entry_number], judge_id]
                    if judge_id in presentation_station_entry_rows[station_id]['entries'][entry_number]:
                        # row.extend([presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][0], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][1], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][2], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][3], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][4], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][5]])
                        row.extend(presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                        if judge_id not in running_totals:
                            running_totals[judge_id] = []
                        running_totals[judge_id].append(presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                    else:
                        row.extend(['','','','','',''])
                    if debugit: print(','.join([str(x) for x in row]))
                    print(','.join([str(x) for x in row]), file=f)
                    last_row = append_row_2(presentation_sheet, row, data_cell_format)
                presentation_sheet.conditional_format(last_entry_row, 2, last_row-1, 2, {'type': '3_color_scale'})
                presentation_sheet.conditional_format(last_entry_row, 3, last_row-1, 3, {'type': '3_color_scale'})
                presentation_sheet.conditional_format(last_entry_row, 4, last_row-1, 4, {'type': '3_color_scale'})
                presentation_sheet.conditional_format(last_entry_row, 5, last_row-1, 5, {'type': '3_color_scale'})
                presentation_sheet.conditional_format(last_entry_row, 6, last_row-1, 6, {'type': '3_color_scale'})
                presentation_sheet.conditional_format(last_entry_row, 7, last_row-1, 7, {'type': '3_color_scale'})
                last_entry_row = last_row
            presentation_sheet.conditional_format(header_row, 8, last_row-1, 12, {'type': '3_color_scale'})
            presentation_sheet.conditional_format(header_row, 13, last_row-1, 13, {'type': '3_color_scale'})
            if debugit: print()
            print('', file=f)
            append_row_2(presentation_sheet, [], data_cell_format)

        if debugit: print("Difficulty\n")
        print("Difficulty\n", file=f)
        all_scores_station_entry_rows = sr_scores_station_entry_rows | dd_scores_station_entry_rows
        for station_id in all_scores_station_entry_rows:
            if station_id == '0000': continue
            judge_scores = {}
            judge_sorted_scores = {}
            judge_scores_ranked = {}
            for judge_type_id in all_scores_station_entry_rows[station_id]['judge_type']:
                if debugit: print(station_id+ ' ' + judge_type_id)
                print(station_id + ' ' + judge_type_id + ' Avg Diff vs Judge Scores', file=f)
                append_row_2(difficulty_sheet, [station_id + ' ' + judge_type_id + ' Avg Diff vs Judge Scores'], bold_cell_format)

                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats']= dict(sorted(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'].items()))
                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'].sort()
                d_avg = {}
                for entry_number in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list']:
                    if len(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][entry_number]) > 0:
                        d_avg[entry_number] = round(sum(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][entry_number])/len(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][entry_number]),2)
                    else:
                        d_avg[entry_number] = 0
                sorted_d_avg = sorted(d_avg.items(), key=lambda x: x[1], reverse=True)

                row = ['Entry Number', 'Davg']
                for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                    row.append(judge_id + ' ' + judge_type_id)
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                header_row = append_row_2(difficulty_sheet, row, data_cell_format)
                num_columns = len(row)
                for entry_number, d in sorted_d_avg:
                    if args.anonymous:
                        my_entry_number = 'Entry n'
                    else:
                        my_entry_number = entry_number
                    row = [my_entry_number + ' ' + all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][entry_number] + ' ' + entry_to_teamname[entry_number], d]
                    for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                        if judge_id not in judge_scores:
                            judge_scores[judge_id] = {}
                            judge_scores_ranked[judge_id] = {}
                        if judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number]:
                            row.append(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number][judge_id][0])
                            judge_scores[judge_id][entry_number] = all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number][judge_id][0]
                        else:
                            row.append('')
                            judge_scores[judge_id][entry_number] = 0
                    if debugit: print(','.join([str(x) for x in row]))
                    print(','.join([str(x) for x in row]), file=f)
                    last_row = append_row_2(difficulty_sheet, row, data_cell_format)
                difficulty_sheet.conditional_format(header_row, 1, last_row-1, num_columns-1, {'type': '3_color_scale'})
                if debugit: print()
                print('', file=f)
                append_row_2(difficulty_sheet, [], data_cell_format)

                if debugit: print(station_id + ' ' + judge_type_id)
                print(station_id + ' ' + judge_type_id + ' Relative Rankings', file=f)
                append_row_2(difficulty_sheet, [station_id + ' ' + judge_type_id + ' Relative Rankings'], bold_cell_format)

                row = ['Entry Number', 'Rank']
                for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                    row.append(judge_id + ' ' + judge_type_id)
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                header_row = append_row_2(difficulty_sheet, row, data_cell_format)
                num_columns = len(row)

                for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                    judge_sorted_scores[judge_id] = dict(sorted(judge_scores[judge_id].items(), key=lambda item: item[1], reverse=True))
                    rank = 1
                    for entry_number in judge_sorted_scores[judge_id]:
                        judge_scores_ranked[judge_id][entry_number] = rank
                        rank += 1
                rank = 1
                for entry_number, d in sorted_d_avg:
                    if args.anonymous:
                        my_entry_number = 'Entry n'
                    else:
                        my_entry_number = entry_number
                    row = [my_entry_number + ' ' + all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][entry_number] + ' ' + entry_to_teamname[entry_number], rank]
                    for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                        if judge_id in judge_scores_ranked:
                            row.append(judge_scores_ranked[judge_id][entry_number])
                        else:
                            row.append('')
                    if debugit: print(','.join([str(x) for x in row]))
                    print(','.join([str(x) for x in row]), file=f)
                    last_row = append_row_2(difficulty_sheet, row, data_cell_format)
                    rank += 1
                difficulty_sheet.conditional_format(header_row, 1, last_row-1, num_columns-1, {'type': 'data_bar'})
                if debugit: print()
                print('', file=f)
                append_row_2(difficulty_sheet, [], data_cell_format)

                append_row_2(difficulty_sheet, [station_id + ' ' + judge_type_id + ' Cummulative scores across all heats'], bold_cell_format)
                row = ['Judge Info']
                row.extend(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'])
                row.append('Heat Count')
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                header_row = append_row_2(difficulty_sheet, row, data_cell_format)
                num_columns = len(row)
                for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats']:
                    row = [judge_id + ' ' + judge_type_id]
                    if all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]['heat_count'] > 0:
                        all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]['d'] = round(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]['d']/all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]['heat_count'],2)
                        all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]['Avg Clicks/Heat'] = round(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]['Avg Clicks/Heat']/all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]['heat_count'],2)
                    row.extend([all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][key] for key in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id]])
                    if debugit: print(','.join([str(x) for x in row]))
                    print(','.join([str(x) for x in row]), file=f)
                    last_row = append_row_2(difficulty_sheet, row, data_cell_format)
                difficulty_sheet.conditional_format(header_row, 2, last_row-1, num_columns-3, {'type': 'data_bar', 'min_value': 0, 'min_type': 'num', })
                difficulty_sheet.conditional_format(header_row, 1, last_row-1, 1, {'type': 'data_bar', 'min_value': 0, 'min_type': 'num', })
                difficulty_sheet.conditional_format(header_row, num_columns-2, last_row-1, num_columns-2, {'type': 'data_bar', 'min_value': 0, 'min_type': 'num', })

                if debugit: print()
                print('', file=f)
                append_row_2(difficulty_sheet, [], data_cell_format)

                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'].sort()

                append_row_2(difficulty_sheet, [station_id + ' ' + judge_type_id + ' Score Details'], bold_cell_format)
                row = ['Entry Number', 'Judge Info']
                row.extend(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'])
                row[-1] = 'Total Clicks'
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                header_row = append_row_2(difficulty_sheet, row, data_cell_format)
                num_columns = len(row)

                for entry_number in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries']:
                    for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                        if args.anonymous:
                            my_entry_number = 'Entry n'
                        else:
                            my_entry_number = entry_number
                        row = [my_entry_number + ' ' + all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][entry_number] + ' ' + entry_to_teamname[entry_number]]
                        row.append(judge_id + ' ' + judge_type_id)
                        if judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number]:
                            row.extend(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number][judge_id])
                        else:
                            row.append('')
                        if debugit: print(','.join([str(x) for x in row]))
                        print(','.join([str(x) for x in row]), file=f)
                        last_row = append_row_2(difficulty_sheet, row, data_cell_format)
                difficulty_sheet.conditional_format(header_row, 2, last_row-1, num_columns-2, {'type': '3_color_scale'})
                difficulty_sheet.conditional_format(header_row, 1, last_row-1, 1, {'type': '3_color_scale'})
                difficulty_sheet.conditional_format(header_row, num_columns-1, last_row-1, num_columns-1, {'type': '3_color_scale'})

                if debugit: print()
                print('', file=f)
                append_row_2(difficulty_sheet, [], data_cell_format)

    set_column_widths(speed_sheet)
    set_column_widths(miss_sheet)
    set_column_widths(break_sheet)
    set_column_widths(presentation_sheet)
    set_column_widths(difficulty_sheet)
    wb.filename = filename.replace(' ', '_') + '-analysis.xlsx'
    wb.close()
    print("Done")
    input('press enter to quit')
    
if __name__ == '__main__':
    main()
