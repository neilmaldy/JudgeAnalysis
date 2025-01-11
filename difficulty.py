import json
import csv
import pprint

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

sr_scores_station_entry_rows = {}
dd_scores_station_entry_rows = {}
for judge_type_id in ['Dr', 'Dm', 'Dp', 'Db', 'Da', 'Dj', 'Dt']:
    for entry_number in scores[judge_type_id]: # Dr Dm Dp Db Da Dj Dt
        print(entry_number)
        for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(scores[judge_type_id][entry_number], key=lambda x: x[1]):
            station_id = judge_id.split('-')[0]
            if event_definition_abbr in ['SRIF', 'SRPF', 'SRTF', 'WHPF'] and judge_type_id in ['Dr', 'Dm', 'Dp', 'Db', 'Da']:
                if station_id not in sr_scores_station_entry_rows:
                    sr_scores_station_entry_rows[station_id] = {}
                    sr_scores_station_entry_rows[station_id]['judge_ids'] = []
                    sr_scores_station_entry_rows[station_id]['judge_types'] = {}
                    sr_scores_station_entry_rows[station_id]['entries'] = {}
                    sr_scores_station_entry_rows[station_id]['entry_types'] = {}
                    sr_scores_station_entry_rows[station_id]['judge_stats'] = {}
                if judge_id not in sr_scores_station_entry_rows[station_id]['judge_ids']:
                    sr_scores_station_entry_rows[station_id]['judge_ids'].append(judge_id)
                if judge_type_id not in sr_scores_station_entry_rows[station_id]['judge_types']:
                    sr_scores_station_entry_rows[station_id]['judge_types'][judge_id] = judge_type_id
                if entry_number not in sr_scores_station_entry_rows[station_id]['entries']:
                    sr_scores_station_entry_rows[station_id]['entries'][entry_number] = {}
                if entry_number not in sr_scores_station_entry_rows[station_id]['entry_types']:
                    sr_scores_station_entry_rows[station_id]['entry_types'][entry_number] = event_definition_abbr
                temp_dict = judge_tally_data
                temp_dict.pop('rep', None)
                temp_dict.pop('break', None)
                temp_dict['d'] = round(judge_results['d'],2)
                if 'columns' not in sr_scores_station_entry_rows[station_id]:
                    sr_scores_station_entry_rows[station_id]['columns'] = ','.join(sorted(temp_dict.keys())) + ',Total'
                temp_list = []
                total_score = 0
                for key in sorted(temp_dict.keys()):
                    temp_list.append(str(temp_dict[key]))
                    if 'diff' in key:
                        total_score += temp_dict[key]
                temp_dict['Total'] = total_score
                temp_list.append(str(total_score))
                sr_scores_station_entry_rows[station_id]['entries'][entry_number][judge_id] = ','.join(temp_list)
                if judge_id not in sr_scores_station_entry_rows[station_id]['judge_stats']:
                    sr_scores_station_entry_rows[station_id]['judge_stats'][judge_id] = {}
                    for key in sr_scores_station_entry_rows[station_id]['columns'].split(','):
                        sr_scores_station_entry_rows[station_id]['judge_stats'][judge_id][key] = 0
                    sr_scores_station_entry_rows[station_id]['judge_stats'][judge_id]['heat_count'] = 0  
                if total_score > 0:
                    for key in sr_scores_station_entry_rows[station_id]['judge_stats'][judge_id]:
                        if key == 'heat_count':
                            sr_scores_station_entry_rows[station_id]['judge_stats'][judge_id][key] += 1
                        else:
                            sr_scores_station_entry_rows[station_id]['judge_stats'][judge_id][key] += temp_dict[key]

            if event_definition_abbr in ['DDSF', 'DDPF'] and judge_type_id in ['Dj', 'Dt']:
                if station_id not in dd_scores_station_entry_rows:
                    dd_scores_station_entry_rows[station_id] = {}
                    dd_scores_station_entry_rows[station_id]['judge_ids'] = []
                    dd_scores_station_entry_rows[station_id]['judge_types'] = {}
                    dd_scores_station_entry_rows[station_id]['entries'] = {}
                    dd_scores_station_entry_rows[station_id]['entry_types'] = {}
                    dd_scores_station_entry_rows[station_id]['judge_stats'] = {}
                if judge_id not in dd_scores_station_entry_rows[station_id]['judge_ids']:
                    dd_scores_station_entry_rows[station_id]['judge_ids'].append(judge_id)
                if judge_type_id not in dd_scores_station_entry_rows[station_id]['judge_types']:
                    dd_scores_station_entry_rows[station_id]['judge_types'][judge_id] = judge_type_id
                if entry_number not in dd_scores_station_entry_rows[station_id]['entries']:
                    dd_scores_station_entry_rows[station_id]['entries'][entry_number] = {}
                if entry_number not in dd_scores_station_entry_rows[station_id]['entry_types']:
                    dd_scores_station_entry_rows[station_id]['entry_types'][entry_number] = event_definition_abbr
                temp_dict = {'d': round(judge_results['d'],2)}
                temp_dict.update(judge_tally_data)
                temp_dict.pop('rep', None)
                temp_dict.pop('break', None)
                temp_dict['d'] = round(judge_results['d'],2)
                if 'columns' not in dd_scores_station_entry_rows[station_id]:
                    dd_scores_station_entry_rows[station_id]['columns'] = ','.join(temp_dict.keys()) + ',Total'
                temp_list = []
                total_score = 0
                for key in sorted(temp_dict.keys()):
                    temp_list.append(str(temp_dict[key]))
                    if 'diff' in key:
                        total_score += temp_dict[key]
                temp_dict['Total'] = total_score
                temp_list.append(str(total_score))
                dd_scores_station_entry_rows[station_id]['entries'][entry_number][judge_id] = ','.join(temp_list)
                if judge_id not in dd_scores_station_entry_rows[station_id]['judge_stats']:
                    dd_scores_station_entry_rows[station_id]['judge_stats'][judge_id] = {}
                    for key in dd_scores_station_entry_rows[station_id]['columns'].split(','):
                        dd_scores_station_entry_rows[station_id]['judge_stats'][judge_id][key] = 0
                    dd_scores_station_entry_rows[station_id]['judge_stats'][judge_id]['heat_count'] = 0  
                if total_score > 0:
                    for key in dd_scores_station_entry_rows[station_id]['judge_stats'][judge_id]:
                        if key == 'heat_count':
                            dd_scores_station_entry_rows[station_id]['judge_stats'][judge_id][key] += 1
                        else:
                            dd_scores_station_entry_rows[station_id]['judge_stats'][judge_id][key] += temp_dict[key]

print()

with open('output.csv', 'w') as f:
    print("Difficulty\n")
    print("Difficulty\n", file=f)
    all_scores_station_entry_rows = sr_scores_station_entry_rows | dd_scores_station_entry_rows
    for station_id in all_scores_station_entry_rows:
        if station_id in station_id_to_session_name:
            print(station_id_to_session_name[station_id])
            print(station_id_to_session_name[station_id], file=f)
        else:
            print(station_id)
            print(station_id, file=f)

        row = 'Judge Info,' + all_scores_station_entry_rows[station_id]['columns'] + ',Heat Count'
        print(row)
        print(row, file=f)
        all_scores_station_entry_rows[station_id]['judge_stats']= dict(sorted(all_scores_station_entry_rows[station_id]['judge_stats'].items()))
        for judge_id in all_scores_station_entry_rows[station_id]['judge_stats']:
            row = judge_id + ' ' + all_scores_station_entry_rows[station_id]['judge_types'][judge_id]
            if all_scores_station_entry_rows[station_id]['judge_stats'][judge_id]['heat_count'] > 0:
                all_scores_station_entry_rows[station_id]['judge_stats'][judge_id]['d'] = round(all_scores_station_entry_rows[station_id]['judge_stats'][judge_id]['d']/all_scores_station_entry_rows[station_id]['judge_stats'][judge_id]['heat_count'],2)
                all_scores_station_entry_rows[station_id]['judge_stats'][judge_id]['Total'] = round(all_scores_station_entry_rows[station_id]['judge_stats'][judge_id]['Total']/all_scores_station_entry_rows[station_id]['judge_stats'][judge_id]['heat_count'],2)
            row += ',' + ','.join([str(all_scores_station_entry_rows[station_id]['judge_stats'][judge_id][key]) for key in all_scores_station_entry_rows[station_id]['judge_stats'][judge_id]])
            print(row)
            print(row, file=f)
        else:
            print()
            print('', file=f)

        all_scores_station_entry_rows[station_id]['judge_ids'].sort()

        row = 'Entry Number,Judge Info,' + all_scores_station_entry_rows[station_id]['columns']
        print(row)
        print(row, file=f)

        for entry_number in all_scores_station_entry_rows[station_id]['entries']:
            for judge_id in all_scores_station_entry_rows[station_id]['judge_ids']:
                row = entry_number + ' ' + all_scores_station_entry_rows[station_id]['entry_types'][entry_number]
                row += ',' + judge_id + ' ' + all_scores_station_entry_rows[station_id]['judge_types'][judge_id]
                if judge_id in all_scores_station_entry_rows[station_id]['entries'][entry_number]:
                    row += ',' + str(all_scores_station_entry_rows[station_id]['entries'][entry_number][judge_id])
                else:
                    row += ','
                print(row)
                print(row, file=f)
        else:
            print()
            print('', file=f)

