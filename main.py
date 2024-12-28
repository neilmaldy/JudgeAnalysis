import json
import csv
import pprint

file = open('CompetitionScores_Fast Feet and Freestyle Faceoff_2024-12-25_21-57-42.tsv', 'r')
dict_reader = csv.DictReader(file, delimiter='\t')
data = [row for row in dict_reader]
file.close()
session_name_to_station_id = {}
session_name_to_station_id['Speed-1'] = '3728'
session_name_to_station_id['Speed-2'] = '3729'
session_name_to_station_id['Speed-3'] = '3730'
session_name_to_station_id['Speed-4'] = '3731'
session_name_to_station_id['Speed-5'] = '3732'
session_name_to_station_id['Speed-6'] = '3733'
session_name_to_station_id['Freestyle-1'] = '3734'
session_name_to_station_id['Double Dutch-1'] = '3830'
session_name_to_station_id['Wheel-1'] = '3831'
session_name_to_station_id['SRTF-1'] = '3832'
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
with open('output.csv', 'w') as f:
    for entry_number in scores['P']:
        print(entry_number)
        sum_of_p = 0
        sum_of_nm = 0
        judge_count = 0
        row = entry_number
        for event_definition_abbr, judge_id, judge_tally_data, judge_results in scores['P'][entry_number]:
            sum_of_p += judge_results['p']
            sum_of_nm += judge_results['nm']
            judge_count += 1
        for event_definition_abbr, judge_id, judge_tally_data, judge_results in sorted(scores['P'][entry_number], key=lambda x: x[1]):
            vs_avg_p = round(judge_results['p'] - (sum_of_p - judge_results['p']) / (judge_count - 1), 2)
            vs_avg_nm = round(judge_results['nm'] - (sum_of_nm - judge_results['nm']) / (judge_count - 1), 2)
            row += ',' + ','.join([judge_id, str(round(judge_results['p'], 2)), str(judge_results['nm']), str(vs_avg_p), str(vs_avg_nm)])
        print(row, file=f)
