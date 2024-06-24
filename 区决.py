import random
import time
from collections import defaultdict
from tqdm import tqdm
from openpyxl import Workbook
from openpyxl.styles import numbers

teams = {
    "LGND": 124,
    "MEKRS": 101,
    "VKG": 96,
    "SWQ": 95,
    "XNY": 90,
    "GUILD": 87,
    "DF": 85,
    "MDY": 77,
    "LVI": 76,
    "BCG": 74,
    "WBG": 74,
    "KD": 66,
    "EVOS": 60,
    "LGD": 59,
    "RTG": 57,
    "GS": 56,
    "WMFB": 55,
    "HEROEZ": 51,
    "FD": 48,
    "GEE": 47,
}

rank_points = [25, 21, 18, 16, 15, 14, 13, 12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 2, 1, 0]
rank_points_match = [12, 9, 7, 5, 4, 3, 3, 2, 2, 2, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0]
simulated_results = set()


def simulate_round(teams):
    ranked_results = ()
    round_results = {}

    while ranked_results == () or ranked_results in simulated_results:
        for _ in range(6):
            random_rank = random.sample(list(teams.keys()), 20)

        for rank, team in enumerate(random_rank):
            if team not in round_results:
                round_results[team] = [0, 0]
            if rank <= 2:
                score = rank_points_match[rank] + random.randint(3, 12)
            elif rank <= 5:
                score = rank_points_match[rank] + random.randint(1, 6)
            else:
                score = rank_points_match[rank] + random.randint(0, 3)
            round_results[team][0] = round_results[team][0] + score
            round_results[team][1] = max(score, round_results[team][1])

        ranked_teams = sorted(
            round_results.items(), key=lambda x: (x[1][0], x[1][1]), reverse=True
        )
        ranked_results = tuple(map(lambda t: t[0], ranked_teams))

    simulated_results.add(ranked_results)

    for i, team in enumerate(ranked_teams):
        teams[team[0]] += rank_points[i]

    return ranked_teams


def determine_qualifiers(teams, round_ranking):
    champion = round_ranking[0][0]
    remaining_teams = sorted(teams.items(), key=lambda x: x[1], reverse=True)
    qualifiers = [champion]
    for team, _ in remaining_teams:
        if team != champion and len(qualifiers) < 7:
            qualifiers.append(team)

    return qualifiers


def calculate_qualification_probability(simulations):
    qualification_counts = {key: defaultdict(int) for key in teams.keys()}
    rank_counts = {key: defaultdict(int) for key in teams.keys()}

    with tqdm(total=simulations, desc="Simulating") as progress_bar:
        for _ in range(simulations):
            current_teams = teams.copy()
            round_ranking = simulate_round(current_teams)
            qualifiers = determine_qualifiers(current_teams, round_ranking)

            for rank, (team, _) in enumerate(round_ranking):
                rank_counts[team][rank + 1] += 1
                if team in qualifiers:
                    qualification_counts[team][rank + 1] += 1
            progress_bar.update(1)

    qualification_probabilities = {key: {} for key in teams.keys()}
    for res_team in rank_counts:
        for rank in rank_counts[res_team]:
            qualification_probabilities[res_team][rank] = (
                qualification_counts[res_team][rank] / rank_counts[res_team][rank]
            )

    return qualification_probabilities


random.seed(int(time.time()))
wb = Workbook()
ws = wb.active
first_row = ["队伍", "当前积分", "晋级概率"]
for i in range(20):
    first_row.append(f"第{i + 1}名")
ws.append(first_row)
simulations=10000000
qualification_probabilities = calculate_qualification_probability(simulations)
for team_name in qualification_probabilities.keys():
    row = [team_name, teams[team_name], ""]
    print(f"晋级概率 ({team_name}):")
    for rank, probability in sorted(qualification_probabilities[team_name].items()):
        row.append(probability)
        if rank == 10 or rank == 20:
            print(f"名次 {rank}:\t{probability:.2%}")
        else:
            print(f"名次 {rank}:\t{probability:.2%}", end="\t")
    ws.append(row)
for row in ws.iter_rows(min_row=2, max_row=21, min_col=4, max_col=23):
    for cell in row:
        cell.number_format = numbers.FORMAT_PERCENTAGE_00
wb.save(f"results_{simulations}.xlsx")