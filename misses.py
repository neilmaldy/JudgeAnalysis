import json
import csv
import pprint
import xlsxwriter

def max_column_width(x, y):
    return max(x, len(str(y)))


def append_row(worksheet, list_to_append):
    try:
        worksheet.write_row(worksheet.row_counter, 0, list_to_append)
        worksheet.column_widths = list(map(max_column_width, worksheet.column_widths, list_to_append))
    except AttributeError:
        worksheet.row_counter = 0
        worksheet.column_widths = [len(x) for x in list_to_append]
        worksheet.write_row(worksheet.row_counter, 0, list_to_append)
    worksheet.row_counter += 1
    return


def append_row_2(worksheet, list_to_append, cell_format=None):

    if not cell_format:
        cell_format = data_cell_format
    try:
        worksheet.write_row(worksheet.row_counter, 0, list_to_append, cell_format)
        worksheet.column_widths = list(map(max_column_width, worksheet.column_widths, list_to_append))
    except AttributeError:
        worksheet.row_counter = 0
        worksheet.column_widths = [len(x) for x in list_to_append]
        worksheet.write_row(worksheet.row_counter, 0, list_to_append, cell_format)
    worksheet.row_counter += 1
    return worksheet.row_counter


def set_column_widths(worksheet):
    last_column_id = None
    last_column_width = None
    for column_id, column_width in enumerate(worksheet.column_widths):
        worksheet.set_column(column_id, column_id, min(50, column_width)+3.0)
        last_column_id = column_id
        last_column_width = column_width
    if last_column_id and last_column_width:
        worksheet.set_column(last_column_id, last_column_id, last_column_width + 5.0)
    return

wb = xlsxwriter.Workbook()
miss_sheet = wb.add_worksheet('Misses')
break_sheet = wb.add_worksheet('Breaks')
presentation_sheet = wb.add_worksheet('Presentation')

data_cell_format = wb.add_format({'border': 1})

file = open('CompetitionScores_YMCA Super Skipper Judge Training_2025-01-06_16-53-04.tsv', 'r')
dict_reader = csv.DictReader(file, delimiter='\t')
data = [row for row in dict_reader]
file.close()
session_name_to_station_id = {}
session_name_to_station_id['Single Rope-1'] = '3867'
session_name_to_station_id['Single Rope-2'] = '3868'
session_name_to_station_id['Wheel-1'] = '3869'
session_name_to_station_id['Wheel-2'] = '3870'
session_name_to_station_id['Double Dutch-1'] = '3871'
session_name_to_station_id['Double Dutch-2'] = '3872'
judge_id_to_name = {}
judge_id_to_name['3867-1'] = 'Darcy S'
judge_id_to_name['3867-2'] = 'Stephanie E'
judge_id_to_name['3867-3'] = 'Fiona W'
judge_id_to_name['3867-4'] = 'Kristen M'
judge_id_to_name['3867-5'] = 'JD D'
judge_id_to_name['3869-1'] = 'Darcy S'
judge_id_to_name['3869-2'] = 'Stephanie E'
judge_id_to_name['3869-3'] = 'Fiona W'
judge_id_to_name['3869-4'] = 'Kristen M'
judge_id_to_name['3869-5'] = 'JD D'
judge_id_to_name['3871-1'] = 'Darcy S'
judge_id_to_name['3871-2'] = 'Stephanie E'
judge_id_to_name['3871-3'] = 'Fiona W'
judge_id_to_name['3871-4'] = 'Kristen M'
judge_id_to_name['3871-5'] = 'JD D'

judge_id_to_name['3868-1'] = 'Jennifer H'
judge_id_to_name['3868-2'] = 'Cheryl C'
judge_id_to_name['3868-3'] = 'Neha P'
judge_id_to_name['3868-4'] = 'Megan D'
judge_id_to_name['3868-5'] = 'Lainie C'
judge_id_to_name['3870-1'] = 'Jennifer H'
judge_id_to_name['3870-2'] = 'Cheryl C'
judge_id_to_name['3870-3'] = 'Neha P'
judge_id_to_name['3870-4'] = 'Megan D'
judge_id_to_name['3870-5'] = 'Lainie C'
judge_id_to_name['3872-1'] = 'Jennifer H'
judge_id_to_name['3872-2'] = 'Cheryl C'
judge_id_to_name['3872-3'] = 'Neha P'
judge_id_to_name['3872-4'] = 'Megan D'
judge_id_to_name['3872-5'] = 'Lainie C'

judge_id_to_name['3867-11'] = 'Will A'
judge_id_to_name['3867-12'] = 'Cynthia M'
judge_id_to_name['3867-13'] = 'Matt A'
judge_id_to_name['3867-14'] = 'Teresa A'
judge_id_to_name['3869-11'] = 'Will A'
judge_id_to_name['3869-12'] = 'Cynthia M'
judge_id_to_name['3869-13'] = 'Matt A'
judge_id_to_name['3869-14'] = 'Teresa A'
judge_id_to_name['3871-11'] = 'Will A'
judge_id_to_name['3871-12'] = 'Cynthia M'
judge_id_to_name['3871-13'] = 'Matt A'
judge_id_to_name['3871-14'] = 'Teresa A'

judge_id_to_name['3868-11'] = 'Heidi B'
judge_id_to_name['3868-12'] = 'Mencken D'
judge_id_to_name['3868-13'] = 'Justin K'
judge_id_to_name['3868-14'] = 'Kenji N'
judge_id_to_name['3870-11'] = 'Heidi B'
judge_id_to_name['3870-12'] = 'Mencken D'
judge_id_to_name['3870-13'] = 'Justin K'
judge_id_to_name['3870-14'] = 'Kenji N'
judge_id_to_name['3872-11'] = 'Heidi B'
judge_id_to_name['3872-12'] = 'Mencken D'
judge_id_to_name['3872-13'] = 'Justin K'
judge_id_to_name['3872-14'] = 'Kenji N'

station_id_to_session_name = {}
for session_name in session_name_to_station_id:
    station_id_to_session_name[session_name_to_station_id[session_name]] = session_name

scores = {}
for row in data:
    try:
        judgedata = row['JudgeScoreDataString']
        competition_name = row['CompetitionName']
        session_name = row['SessionName']
        session_id = row['SessionID']
        entry_number = row['EntryNumber']
        event_definition_abbr = row['EventDefinitionAbbr']
        station_id = str(row['StationID'])
        station_sequence = str(row['StationSequence'])
        score_sequence = str(row['ScoreSequence'])
        if not station_id:
            if session_name + '-' + station_sequence in session_name_to_station_id:
                station_id = session_name_to_station_id[ + '-' + station_sequence]
            else:
                station_id = '0000'
        judge_id = station_id + '-' + score_sequence
        is_scored = row['IsScored']
        total_score = row['TotalScore']
        entry_is_scored = row['EntryIsScored']
        is_locked = row['IsLocked']
        judge_is_scored = row['JudgeIsScored']
        if is_scored == 'True':
            judge_score_data = json.loads(row['JudgeScoreDataString'])
            judge_meta_data = judge_score_data['JudgeResults']['meta']
            judge_tally_data = judge_score_data['TallySheet']['tally']
            judge_results = judge_score_data['JudgeResults']['result']
            if judge_meta_data['judgeTypeId'] not in scores:
                scores[judge_meta_data['judgeTypeId']] = {}
            if entry_number not in scores[judge_meta_data['judgeTypeId']]:
                scores[judge_meta_data['judgeTypeId']][entry_number] = []
            scores[judge_meta_data['judgeTypeId']][entry_number].append((event_definition_abbr, judge_id, judge_tally_data, judge_results))
            # pprint.pprint(judge_tally_data)
            # pprint.pprint(judge_results)
    except Exception as e:
        print(e)
        print("Problem with entry number: ", entry_number)

print("Scores parsed")

for judge_type_id in scores:
    for entry_number in scores[judge_type_id]:
        print(judge_type_id, entry_number)
        for event_definition_abbr, judge_id, judge_tally_data, judge_results in scores[judge_type_id][entry_number]:
            print(event_definition_abbr, judge_id, judge_tally_data)
            # pprint.pprint(judge_tally_data)
            pprint.pprint(judge_results)

misses_station_entry_rows = {}
presentation_station_entry_rows = {}
breaks_station_entry_rows = {}

for entry_number in scores['P']:
    print(entry_number)
    for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(scores['P'][entry_number], key=lambda x: x[1]):
        station_id = judge_id.split('-')[0]
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
        presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id] = (round(judge_results['p'], 2), judge_tally_data['ent'], judge_tally_data['form'], judge_tally_data['music'], judge_tally_data['crea'], judge_tally_data['vari'])
        presentation_station_entry_rows[station_id]['entries'][entry_number]['p_list'].append(round(judge_results['p'], 2))
        if judge_id not in presentation_station_entry_rows[station_id]['judge_stats']:
            presentation_station_entry_rows[station_id]['judge_stats'][judge_id] = []
        presentation_station_entry_rows[station_id]['judge_stats'][judge_id].append((round(judge_results['p'], 2), judge_tally_data['ent'], judge_tally_data['form'], judge_tally_data['music'], judge_tally_data['crea'], judge_tally_data['vari']))
for entry_number in scores['T']:
    print(entry_number)
    for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(scores['T'][entry_number], key=lambda x: x[1]):
        station_id = judge_id.split('-')[0]
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
    print(entry_number)
    for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(scores['Dj'][entry_number], key=lambda x: x[1]):
        station_id = judge_id.split('-')[0]
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


with open('output.csv', 'w') as f:
    print("Misses\n")
    print("Misses\n", file=f)
    for station_id in misses_station_entry_rows:
        print(station_id)
        print(station_id, file=f)
        current_row = append_row_2(miss_sheet, [station_id], data_cell_format)
        misses_station_entry_rows[station_id]['judge_ids'].sort()
        row = ['Entry Number']
        for judge_id in misses_station_entry_rows[station_id]['judge_ids']:
            row.append(judge_id + ' ' + misses_station_entry_rows[station_id]['judge_types'][judge_id])
        # row = 'Entry Number,' + ','.join(station_entry_rows[station_id]['judge_ids'])
        print(','.join([str(x) for x in row]))
        print(','.join([str(x) for x in row]), file=f)
        header_row = append_row_2(miss_sheet, row, data_cell_format)
        num_columns = len(row)

        running_totals = {}

        # change row strings to lists
        for entry_number in misses_station_entry_rows[station_id]['entries']:
            row = [entry_number + ' ' + misses_station_entry_rows[station_id]['entry_types'][entry_number]]
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
            print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            last_row = append_row_2(miss_sheet, row, data_cell_format)
        miss_sheet.conditional_format(header_row-1, 1, last_row-1, num_columns-1, {'type': '3_color_scale'})
        row = ['Totals']
        for judge_id in misses_station_entry_rows[station_id]['judge_ids']:
            if judge_id in running_totals:
                row.append(sum(running_totals[judge_id]))
            else:
                row.append('')
        print(','.join([str(x) for x in row]))
        print(','.join([str(x) for x in row]), file=f)
        append_row_2(miss_sheet, row, data_cell_format)

        row = ['Averages']
        for judge_id in misses_station_entry_rows[station_id]['judge_ids']:
            if judge_id in running_totals:
                row.append(round(sum(running_totals[judge_id])/len(running_totals[judge_id]), 2))
            else:
                row.append('')
        print(','.join([str(x) for x in row]))
        print(','.join([str(x) for x in row]), file=f)
        append_row_2(miss_sheet, row, data_cell_format)
        print()
        print('', file=f)
        append_row_2(miss_sheet, [], data_cell_format)

    print("Breaks\n")
    print("Breaks\n", file=f)
    for station_id in breaks_station_entry_rows:
        print(station_id)
        print(station_id, file=f)
        append_row_2(break_sheet, [station_id], data_cell_format)

        breaks_station_entry_rows[station_id]['judge_ids'].sort()
        row = ['Entry Number']
        for judge_id in breaks_station_entry_rows[station_id]['judge_ids']:
            row.append(judge_id + ' ' + breaks_station_entry_rows[station_id]['judge_types'][judge_id])
        # row = 'Entry Number,' + ','.join(station_entry_rows[station_id]['judge_ids'])
        print(','.join([str(x) for x in row]))
        print(','.join([str(x) for x in row]), file=f)
        append_row_2(break_sheet, row, data_cell_format)

        for entry_number in breaks_station_entry_rows[station_id]['entries']:
            row = [entry_number + ' ' + breaks_station_entry_rows[station_id]['entry_types'][entry_number]]
            for judge_id in breaks_station_entry_rows[station_id]['judge_ids']:
                if judge_id in breaks_station_entry_rows[station_id]['entries'][entry_number]:
                    row.append(breaks_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                    if judge_id not in running_totals:
                        running_totals[judge_id] = []
                    running_totals[judge_id].append(breaks_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                else:
                    row.append('')
            print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            append_row_2(break_sheet, row, data_cell_format)

        row = ['Totals']
        for judge_id in breaks_station_entry_rows[station_id]['judge_ids']:
            if judge_id in running_totals:
                row.append(sum(running_totals[judge_id]))
            else:
                row.append('')
        print(','.join([str(x) for x in row]))
        print(','.join([str(x) for x in row]), file=f)
        append_row_2(break_sheet, row, data_cell_format)

        row = ['Averages']
        for judge_id in breaks_station_entry_rows[station_id]['judge_ids']:
            if judge_id in running_totals:
                row.append(round(sum(running_totals[judge_id])/len(running_totals[judge_id]), 2))
            else:
                row.append('')
        print(','.join([str(x) for x in row]))
        print(','.join([str(x) for x in row]), file=f)
        append_row_2(break_sheet, row, data_cell_format)

        print()
        print('', file=f)
        append_row_2(break_sheet, [], data_cell_format)

    print("Presentation\n")
    print("Presentation\n", file=f)

    running_totals = {}
    for station_id in misses_station_entry_rows:
        print(station_id)
        print(station_id, file=f)
        append_row_2(presentation_sheet, [station_id], data_cell_format)
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
        print(','.join([str(x) for x in row]))
        print(','.join([str(x) for x in row]), file=f)
        append_row_2(presentation_sheet, row, data_cell_format)

        for entry_number, p_value in sorted_p_avg:
            row = [entry_number + ' ' + presentation_station_entry_rows[station_id]['entry_types'][entry_number], p_avg[entry_number]]
            for judge_id in presentation_station_entry_rows[station_id]['judge_ids']:
                if judge_id in presentation_station_entry_rows[station_id]['entries'][entry_number]:
                    row.append(presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][0])
                else:
                    row.append('')
            print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            append_row_2(presentation_sheet, row, data_cell_format)
        print()
        print('', file=f)
        append_row_2(presentation_sheet, [], data_cell_format)

        row = ['Judge', 'P avg', 'E avg', 'F avg', 'M avg', 'C avg', 'V avg']
        print(','.join([str(x) for x in row]))
        print(','.join([str(x) for x in row]), file=f)
        append_row_2(presentation_sheet, row, data_cell_format)

        for judge_id in presentation_station_entry_rows[station_id]['judge_ids']:
            p_list = [x[0] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
            e_list = [x[1] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
            f_list = [x[2] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
            m_list = [x[3] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
            c_list = [x[4] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
            v_list = [x[5] for x in presentation_station_entry_rows[station_id]['judge_stats'][judge_id]]
            row = [judge_id]
            row.extend([round(sum(p_list)/len(p_list), 2), round(sum(e_list)/len(e_list), 2), round(sum(f_list)/len(f_list), 2), round(sum(m_list)/len(m_list), 2), round(sum(c_list)/len(c_list), 2), round(sum(v_list)/len(v_list), 2)])
            print(','.join([str(x) for x in row]))
            print(','.join([str(x) for x in row]), file=f)
            append_row_2(presentation_sheet, row, data_cell_format)
        print()
        print('', file=f)
        append_row_2(presentation_sheet, [], data_cell_format)

        row = ['Entry Number', 'Judge', 'P', 'E', 'F', 'M', 'C', 'V']
        print(','.join([str(x) for x in row]))
        print(','.join([str(x) for x in row]), file=f)
        append_row_2(presentation_sheet, row, data_cell_format)

        for entry_number in presentation_station_entry_rows[station_id]['entries']:
            for judge_id in presentation_station_entry_rows[station_id]['judge_ids']:
                row = [entry_number + ' ' + presentation_station_entry_rows[station_id]['entry_types'][entry_number], judge_id]
                if judge_id in presentation_station_entry_rows[station_id]['entries'][entry_number]:
                    row.extend([presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][0], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][1], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][2], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][3], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][4], presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id][5]])
                    if judge_id not in running_totals:
                        running_totals[judge_id] = []
                    running_totals[judge_id].append(presentation_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                else:
                    row.append('','','','','','')
                print(','.join([str(x) for x in row]))
                print(','.join([str(x) for x in row]), file=f)
                append_row_2(presentation_sheet, row, data_cell_format)

        print()
        print('', file=f)
        append_row_2(presentation_sheet, [], data_cell_format)

set_column_widths(miss_sheet)
set_column_widths(break_sheet)
set_column_widths(presentation_sheet)
wb.filename = 'output.xlsx'
wb.close()
print("Done")
