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
            worksheet.set_column(column_id, column_id, min(20, column_width * 0.85) + 1.0)
            last_column_id = column_id
            last_column_width = column_width
        if False and last_column_id and last_column_width:
            worksheet.set_column(last_column_id, last_column_id, last_column_width)
    return


def which_judge_is_dropped(judges_and_scores=None):
    if not judges_and_scores:
        return None
    if len(judges_and_scores) != 3:
        return None
    sorted_judges = sorted(judges_and_scores, key=lambda x: x[1])
    low_judge_id, low_score = sorted_judges[0]
    mid_judge_id, mid_score = sorted_judges[1]
    high_judge_id, high_score = sorted_judges[2]
    if (high_score - mid_score) > (mid_score - low_score):
        return high_judge_id
    else:
        return low_judge_id


def main():
    parser = argparse.ArgumentParser(description='analyze_scores.py')

    debugit = False
    wb = xlsxwriter.Workbook()
    miss_sheet = wb.add_worksheet('Misses')
    break_sheet = wb.add_worksheet('Breaks')
    difficulty_sheet = wb.add_worksheet('Difficulty')

    data_cell_format = wb.add_format({'border': 1})
    bold_cell_format = wb.add_format({'bold': True})
    blue_bg_cell_format = wb.add_format({'bg_color': '#CCE5FF', 'border': 1})
    one_decimal_format = wb.add_format({'num_format': '0.0'})
    two_decimal_format = wb.add_format({'num_format': 2})
    percent_format = wb.add_format({'num_format': '0%'})
    parser.add_argument('filename', metavar='filename', type=str, nargs='?', default='', help='Scoring file name')
    parser.add_argument('-a', '--anonymous', help='Do not include entry numbers', action='store_true')
    try:
        args = parser.parse_args()
        args.filename = r'C:\NoBackup\VSCode_Projects\JudgeAnalysis\scores.tsv'
        if args.filename:
            print("args.filename: ", args.filename)
            sleep(0.2)
            dirname = path.dirname(args.filename)
            if dirname:
                chdir(dirname)
            filename = path.basename(args.filename)
            print("Filename: ", filename)
        else:
            # filename = 'CompetitionScores_Australian Rope Skipping Championship 2025_2025-06-17_22-09-33.tsv'
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
    if not args.anonymous and path.exists('entries.csv'):
        try:
            print("Reading entries.csv")
            sleep(0.2)
            # input('press enter')
            file = open('entries.csv', 'r', encoding='utf-8')
            dict_reader = csv.DictReader(file)
            for row in dict_reader:
                entry_to_teamname[row['EntryNumber']] = row['TeamName'] + ' s/h:' + row['StationSequence'] + '/' + \
                                                        row['HeatNumber'].split('.')[0] + ' r:' + str(row['Rank'])
            file.close()
            print("entries.csv read")
            sleep(0.2)
            # input('press enter')
        except Exception as e:
            print(str(e))
            print("Problem reading entries.csv")
            sleep(0.2)
    elif args.anonymous:
        print("anonymous flag set, team names will not be included")
    else:
        print("entries.csv not found, team names will not be included")

    if not args.anonymous and path.exists('judges.tsv'):
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
    elif args.anonymous:
        print("anonymous flag set, judge names will not be included")
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
            event_definition_abbr = row['EventDefinitionAbbr'] + '_' + row['GenderAbbr'] + '_' + row['AgeGroupName']
            if row['EventDefinitionAbbr'] in ['DDCF', 'SCTF']:
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
            if is_scored == 'True' or True:
                judge_score_data = json.loads(row['JudgeScoreDataString'])
                if 'JudgeResults' not in judge_score_data:
                    print("No judge results for entry number: " + entry_number + " judge_id: " + judge_id)
                    continue
                # todo check for DDCF and TeamShow
                judge_meta_data = judge_score_data['JudgeResults']['meta']
                judge_tally_data = judge_score_data['TallySheet']['tally']
                if judge_meta_data['judgeTypeId'] not in ['Dr', 'Dm', 'Dp', 'Db', 'Da', 'Dj', 'Dt']:
                    continue
                judge_results = judge_score_data['JudgeResults']['result']
                if judge_meta_data['judgeTypeId'] not in scores:
                    scores[judge_meta_data['judgeTypeId']] = {}
                if entry_number not in scores[judge_meta_data['judgeTypeId']]:
                    scores[judge_meta_data['judgeTypeId']][entry_number] = []
                scores[judge_meta_data['judgeTypeId']][entry_number].append(
                    (event_definition_abbr, judge_id, judge_tally_data, judge_results))
                # pprint.pprint(judge_tally_data)
                # pprint.pprint(judge_results)
        except Exception as e:
            print(str(e))
            print("Problem with entry number: ", entry_number)

    print("Scores parsed")
    sleep(0.2)

    for judge_type in ['Dr', 'Dm', 'Dp', 'Db', 'Da', 'Dj', 'Dt']:
        if judge_type not in scores:
            if debugit: print("No data for judge type: ", judge_type)
            scores[judge_type] = {}

    if debugit:
        for judge_type_id in scores:
            for entry_number in scores[judge_type_id]:
                if debugit:
                    print(judge_type_id, entry_number)
                    for event_definition_abbr, judge_id, judge_tally_data, judge_results in scores[judge_type_id][
                        entry_number]:
                        print(event_definition_abbr, judge_id, judge_tally_data)
                        # pprint.pprint(judge_tally_data)
                        pprint.pprint(judge_results)

    sr_scores_station_entry_rows = {}
    dd_scores_station_entry_rows = {}
    for judge_type_id in ['Dr', 'Dm', 'Dp', 'Db', 'Da', 'Dj', 'Dt']:
        for entry_number in scores[judge_type_id]:  # Dr Dm Dp Db Da Dj Dt
            if debugit: print(entry_number)
            for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(
                    scores[judge_type_id][entry_number], key=lambda x: x[1]):
                station_id = judge_id.split('-')[0]
                if station_id == '0000': continue
                if any(event_abbr in event_definition_abbr for event_abbr in
                       ['SRIF', 'SRPF', 'SRTF', 'WHPF']) and judge_type_id in ['Dr', 'Dm', 'Dp', 'Db', 'Da']:
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
                    if judge_id not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'judge_ids']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'].append(
                            judge_id)
                    if entry_number not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'd_list']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][
                            entry_number] = []
                    sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][
                        entry_number].append(round(judge_results['d'], 2))
                    if entry_number not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'entries']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][
                            entry_number] = {}
                    if entry_number not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'entry_types']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][
                            entry_number] = event_definition_abbr
                    temp_dict = judge_tally_data
                    temp_dict.pop('rep', None)
                    temp_dict.pop('break', None)
                    temp_dict['d'] = round(judge_results['d'], 2)
                    if 'columns' not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'] = sorted(
                            temp_dict.keys())
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'].append(
                            'Avg Clicks/Heat')
                    temp_list = []
                    total_score = 0
                    for key in sorted(temp_dict.keys()):
                        temp_list.append(temp_dict[key])
                        if 'diff' in key:
                            total_score += temp_dict[key]
                    temp_dict['Total'] = total_score
                    temp_list.append(total_score)
                    sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number][
                        judge_id] = tuple(temp_list)
                    if judge_id not in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'judge_stats']:
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                            judge_id] = {}
                        for key in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns']:
                            sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                judge_id][key] = 0
                        sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][
                            'heat_count'] = 0
                    if total_score > 0:
                        for key in sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                            judge_id]:
                            if key == 'heat_count':
                                sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                    judge_id][key] += 1
                            elif key == 'Avg Clicks/Heat':
                                sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                    judge_id][key] += temp_dict['Total']
                            else:
                                sr_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                    judge_id][key] += temp_dict[key]

                if any(event_abbr in event_definition_abbr for event_abbr in ['DDSF', 'DDPF']) and judge_type_id in [
                    'Dj', 'Dt']:
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
                    if judge_id not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'judge_ids']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'].append(
                            judge_id)
                    if entry_number not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'd_list']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][
                            entry_number] = []
                    dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][
                        entry_number].append(round(judge_results['d'], 2))
                    if entry_number not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'entries']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][
                            entry_number] = {}
                    if entry_number not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'entry_types']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][
                            entry_number] = event_definition_abbr
                    temp_dict = {'d': round(judge_results['d'], 2)}
                    temp_dict.update(judge_tally_data)
                    temp_dict.pop('rep', None)
                    temp_dict.pop('break', None)
                    temp_dict['d'] = round(judge_results['d'], 2)
                    if 'columns' not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'] = list(
                            temp_dict.keys())
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'].append(
                            'Avg Clicks/Heat')
                    temp_list = []
                    total_score = 0
                    for key in sorted(temp_dict.keys()):
                        temp_list.append(temp_dict[key])
                        if 'diff' in key:
                            total_score += temp_dict[key]
                    temp_dict['Total'] = total_score
                    temp_list.append(total_score)
                    dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][entry_number][
                        judge_id] = tuple(temp_list)
                    if judge_id not in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                        'judge_stats']:
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                            judge_id] = {}
                        for key in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns']:
                            dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                judge_id][key] = 0
                        dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][
                            'heat_count'] = 0
                    if total_score > 0:
                        for key in dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                            judge_id]:
                            if key == 'heat_count':
                                dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                    judge_id][key] += 1
                            elif key == 'Avg Clicks/Heat':
                                dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                    judge_id][key] += temp_dict['Total']
                            else:
                                dd_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                    judge_id][key] += temp_dict[key]

    with open('output.csv', 'w') as f:
        if debugit: print("Difficulty\n")
        print("Difficulty\n", file=f)
        current_chart_row = 1
        all_scores_station_entry_rows = sr_scores_station_entry_rows | dd_scores_station_entry_rows
        for station_id in all_scores_station_entry_rows:
            if station_id == '0000': continue
            judge_scores = {}
            judge_sorted_scores = {}
            judge_scores_ranked = {}
            for judge_type_id in all_scores_station_entry_rows[station_id]['judge_type']:
                if debugit: print(station_id + ' ' + judge_type_id)

                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'] = dict(sorted(
                    all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'].items()))
                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'].sort()

                append_row_2(difficulty_sheet,
                             [station_id + ' ' + judge_type_id + ' Cummulative scores across all heats'],
                             bold_cell_format)
                row = ['Judge Info']
                row.extend(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'])
                row.append('Heat Count')
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                header_row = append_row_2(difficulty_sheet, row, data_cell_format)
                num_columns = len(row)
                for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats']:
                    row = [judge_id + ' ' + judge_type_id]
                    if all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][
                        'heat_count'] > 0:
                        all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][
                            'd'] = round(
                            all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                judge_id]['d'] /
                            all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                judge_id]['heat_count'], 2)
                        all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][judge_id][
                            'Avg Clicks/Heat'] = round(
                            all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                judge_id]['Avg Clicks/Heat'] /
                            all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                judge_id]['heat_count'], 2)
                    row.extend([all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                    judge_id][key] for key in
                                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_stats'][
                                    judge_id]])
                    if debugit: print(','.join([str(x) for x in row]))
                    print(','.join([str(x) for x in row]), file=f)
                    last_row = append_row_2(difficulty_sheet, row, data_cell_format)
                difficulty_sheet.conditional_format(header_row, 2, last_row - 1, num_columns - 3,
                                                    {'type': 'data_bar', 'min_value': 0, 'min_type': 'num', })
                difficulty_sheet.conditional_format(header_row, 1, last_row - 1, 1,
                                                    {'type': 'data_bar', 'min_value': 0, 'min_type': 'num', })
                difficulty_sheet.conditional_format(header_row, num_columns - 2, last_row - 1, num_columns - 2,
                                                    {'type': 'data_bar', 'min_value': 0, 'min_type': 'num', })

                if debugit: print()
                print('', file=f)
                append_row_2(difficulty_sheet, [], data_cell_format)

                print(station_id + ' ' + judge_type_id + ' Avg Diff vs Judge Scores', file=f)
                append_row_2(difficulty_sheet, [station_id + ' ' + judge_type_id + ' Avg Diff vs Judge Scores'],
                             bold_cell_format)

                d_avg = {}
                for entry_number in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list']:
                    if len(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][
                               entry_number]) > 0:
                        d_avg[entry_number] = round(sum(
                            all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][
                                entry_number]) / len(
                            all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['d_list'][
                                entry_number]), 2)
                    else:
                        d_avg[entry_number] = 0
                sorted_d_avg = sorted(d_avg.items(), key=lambda x: x[1], reverse=True)

                row = ['Entry', 'Davg']
                for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                    row.append(judge_id + ' ' + judge_type_id)
                last_judge_column = len(row) - 1

                if last_judge_column == 3:
                    difference_columns = ['E', 'F']
                    count_column = None
                elif last_judge_column == 4:
                    difference_columns = ['F', 'G', 'H']
                    count_column = 'I'
                    # Create a mapping from difference_columns to judge_ids
                    column_to_judge_number = {}
                    if difference_columns:
                        for col, judge_id in zip(difference_columns,
                                                 all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id][
                                                     'judge_ids']):
                            column_to_judge_number[col] = judge_id
                elif last_judge_column == 5:
                    difference_columns = ['G', 'H', 'I', 'J']
                    count_column = None
                else:
                    difference_columns = None
                    count_column = None

                for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                    row.append(judge_id + ' ' + judge_type_id + ' Diff')
                num_columns = len(row)
                if len(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']) == 3:
                    row.append('Dropped Judge')
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                header_row = append_row_2(difficulty_sheet, row, data_cell_format)
                # difficulty_sheet.set_row(header_row-1, None, None, {'level':1, 'hidden': True})
                for entry_number, d in sorted_d_avg:
                    difficulty_scores = []
                    judge_ids_and_judge_scores = []
                    if args.anonymous:
                        my_entry_number = 'Entry n'
                    else:
                        my_entry_number = entry_number
                    row = [my_entry_number + ' ' +
                           all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][
                               entry_number] + ' ' + entry_to_teamname[entry_number], d]
                    for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                        if judge_id not in judge_scores:
                            judge_scores[judge_id] = {}
                            judge_scores_ranked[judge_id] = {}
                        if judge_id in \
                                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][
                                    entry_number]:
                            row.append(
                                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][
                                    entry_number][judge_id][0])
                            judge_scores[judge_id][entry_number] = \
                            all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][
                                entry_number][judge_id][0]
                            difficulty_scores.append(
                                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][
                                    entry_number][judge_id][0])
                            judge_ids_and_judge_scores.append((judge_id,
                                                               all_scores_station_entry_rows[station_id]['judge_type'][
                                                                   judge_type_id]['entries'][entry_number][judge_id][
                                                                   0]))
                        else:
                            row.append('')
                            judge_scores[judge_id][entry_number] = 0
                            difficulty_scores.append('')
                    for difficulty_score in difficulty_scores:
                        if difficulty_score != '':
                            row.append(round(difficulty_score - d, 2))
                        else:
                            row.append('')
                    if len(judge_ids_and_judge_scores) == 3:
                        row.append(which_judge_is_dropped(judge_ids_and_judge_scores))
                    if debugit: print(','.join([str(x) for x in row]))
                    print(','.join([str(x) for x in row]), file=f)
                    last_row = append_row_2(difficulty_sheet, row, data_cell_format)
                    difficulty_sheet.set_row(last_row - 1, None, None, {'level': 1, 'hidden': True})
                difficulty_sheet.conditional_format(header_row, 1, last_row - 1, last_judge_column,
                                                    {'type': '3_color_scale'})
                difficulty_sheet.conditional_format(header_row, last_judge_column + 1, last_row - 1, num_columns - 1,
                                                    {'type': '3_color_scale', 'mid_type': 'num', 'mid_value': 0.0,
                                                     'min_color': 'red', 'mid_color': 'white', 'max_color': 'blue'})
                if difference_columns:
                    d_avg_range = "=Difficulty!$B$" + str(header_row + 1) + ":$B$" + str(last_row)
                    judge_data_ranges = {}
                    judge_id_ranges = {}
                    for column in difference_columns:
                        difficulty_sheet.write_formula(column + str(last_row + 1),
                                                       '=AVERAGE(' + column + str(header_row + 1) + ':' + column + str(
                                                           last_row) + ')', two_decimal_format)
                        difficulty_sheet.write_formula(column + str(last_row + 2),
                                                       '=STDEV(' + column + str(header_row + 1) + ':' + column + str(
                                                           last_row) + ')', two_decimal_format)
                        judge_data_ranges[column] = "=Difficulty!$" + column + "$" + str(
                            header_row + 1) + ":$" + column + "$" + str(last_row) + ""
                        judge_id_ranges[column] = "=Difficulty!$" + column + "$" + str(header_row)
                        if count_column:
                            difficulty_sheet.write_formula(column + str(last_row + 3), '=COUNTIF(' + count_column + str(
                                header_row + 1) + ':' + count_column + str(last_row) + ',"' + str(
                                column_to_judge_number[column]) + '")')
                append_row_2(difficulty_sheet, ['Average Error: '], bold_cell_format)
                append_row_2(difficulty_sheet, ['Stdev: '], bold_cell_format)
                if count_column:
                    append_row_2(difficulty_sheet, ['Drop Count: '], bold_cell_format)
                if debugit: print()
                print('', file=f)
                append_row_2(difficulty_sheet, [], data_cell_format)

                if debugit: print(station_id + ' ' + judge_type_id)
                print(station_id + ' ' + judge_type_id + ' Relative Rankings', file=f)
                append_row_2(difficulty_sheet, [station_id + ' ' + judge_type_id + ' Relative Rankings'],
                             bold_cell_format)

                row = ['Entry', 'Rank']
                for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                    row.append(judge_id + ' ' + judge_type_id)
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                header_row = append_row_2(difficulty_sheet, row, data_cell_format)
                difficulty_sheet.set_row(header_row - 1, None, None, {'level': 1, 'hidden': True})
                num_columns = len(row)

                for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                    judge_sorted_scores[judge_id] = dict(
                        sorted(judge_scores[judge_id].items(), key=lambda item: item[1], reverse=True))
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
                    row = [my_entry_number + ' ' +
                           all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][
                               entry_number] + ' ' + entry_to_teamname[entry_number], rank]
                    for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                        if judge_id in judge_scores_ranked:
                            row.append(judge_scores_ranked[judge_id][entry_number])
                        else:
                            row.append('')
                    if debugit: print(','.join([str(x) for x in row]))
                    print(','.join([str(x) for x in row]), file=f)
                    last_row = append_row_2(difficulty_sheet, row, data_cell_format)
                    difficulty_sheet.set_row(last_row - 1, None, None, {'level': 1, 'hidden': True})
                    rank += 1
                difficulty_sheet.conditional_format(header_row, 1, last_row - 1, num_columns - 1, {'type': 'data_bar'})
                if debugit: print()
                print('', file=f)
                append_row_2(difficulty_sheet, [], data_cell_format)

                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids'].sort()

                append_row_2(difficulty_sheet, [station_id + ' ' + judge_type_id + ' Score Details'], bold_cell_format)
                row = ['Entry', 'Judge Info']
                row.extend(all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['columns'])
                row[-1] = 'Total Clicks'
                if debugit: print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                header_row = append_row_2(difficulty_sheet, row, data_cell_format)
                difficulty_sheet.set_row(header_row - 1, None, None, {'level': 1, 'hidden': True})
                num_columns = len(row)
                last_entry_row = header_row
                for entry_number in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries']:
                    for judge_id in all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['judge_ids']:
                        if args.anonymous:
                            my_entry_number = 'Entry n'
                        else:
                            my_entry_number = entry_number
                        row = [my_entry_number + ' ' +
                               all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entry_types'][
                                   entry_number] + ' ' + entry_to_teamname[entry_number]]
                        row.append(judge_id + ' ' + judge_type_id)
                        if judge_id in \
                                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][
                                    entry_number]:
                            row.extend(
                                all_scores_station_entry_rows[station_id]['judge_type'][judge_type_id]['entries'][
                                    entry_number][judge_id])
                        else:
                            row.append('')
                        if debugit: print(','.join([str(x) for x in row]))
                        print(','.join([str(x) for x in row]), file=f)
                        last_row = append_row_2(difficulty_sheet, row, data_cell_format)
                        difficulty_sheet.set_row(last_row - 1, None, None, {'level': 1, 'hidden': True})
                    difficulty_sheet.conditional_format(last_entry_row, 3, last_row - 1, num_columns - 2,
                                                        {'type': '3_color_scale'})
                    difficulty_sheet.conditional_format(last_entry_row, 2, last_row - 1, 2,
                                                        {'type': '3_color_scale', 'min_type': 'num', 'min_value': 0})
                    last_entry_row = last_row
                # difficulty_sheet.conditional_format(header_row, 3, last_row-1, num_columns-2, {'type': '3_color_scale'})
                # difficulty_sheet.conditional_format(header_row, 2, last_row-1, 2, {'type': '3_color_scale'})
                difficulty_sheet.conditional_format(header_row, num_columns - 1, last_row - 1, num_columns - 1,
                                                    {'type': '3_color_scale'})

                if debugit: print()
                print('', file=f)
                append_row_2(difficulty_sheet, [], data_cell_format)

    set_column_widths(difficulty_sheet)
    wb.filename = filename.replace(' ', '_') + '-analysis.xlsx'
    wb.close()
    print("Done")
    input('press enter to quit')


if __name__ == '__main__':
    main()
