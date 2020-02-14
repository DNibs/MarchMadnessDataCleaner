"""
Cleans data on excel sheet for march madness dataset.
Used for CY305.
"""

import pandas as pd

excel_fn = 'NCAABasicStatsClean.xlsx'
out_fn = 'out.xlsx'
db_master_name = 'NCAAMStats'
db_tourney_name = 'NCAATourney'
db_master = pd.read_excel(excel_fn, sheet_name=db_master_name)
db_tourney = pd.read_excel(excel_fn, sheet_name=db_tourney_name, index_col=None)
# print(db_master.head())
# print(db_tourney.head())

db_out = db_master.copy()
master_count = db_master.shape[0]
tourney_count = db_tourney.shape[0]
print('\n')

# Find teams from tournament database that don't match names from master database
print('Searching for tournament names missing from master database...')
bad_tm_names = []
for j in range(0, tourney_count):
    print("Row {} of {}".format(j, tourney_count))
    team_nm = db_tourney.iloc[j, 3]
    team_in_master = False
    for i in range(0, master_count):
        if db_master.iloc[i, 2] == team_nm:
            team_in_master = True
    if not team_in_master:
        print(team_nm)
        bad_tm_names.append(team_nm)

print(bad_tm_names)
with open('listfile.txt', 'w') as filehandle:
    for listitem in bad_tm_names:
        filehandle.write('%s\n' % listitem)

if len(bad_tm_names) > 0:
    print('Fix names before moving on!')
    exit()

print('Team Names all match.')

print('Merging tournament data and master data into output database file.')
# Fill out tournament results data in output database
for i in range(0, master_count):
    print('Row {} of {}'.format(i, master_count))

    # zero out blank entries
    if pd.isnull(db_out.iloc[i, 4]):
        db_out.iloc[i, 4] = 0
    for k in range(5, 13):
        db_out.iloc[i, k] = False

    # test if team went to tournament before executing code
    if db_master.iloc[i, 3]:
        team_nm = db_master.iloc[i, 2]
        year = db_master.iloc[i, 1]

        # Iterate over every tournament instance if team made tournament
        for j in range(0, tourney_count):
            tour_year = db_tourney.iloc[j, 0]
            tour_ech = db_tourney.iloc[j, 1]
            tour_seed = db_tourney.iloc[j, 2]
            tour_team = db_tourney.iloc[j, 3]

            # Set values in output database for specific year and team
            if year == tour_year:
                if team_nm == tour_team:
                    db_out.iloc[i, 4] = tour_seed
                    db_out.loc[i, tour_ech] = True

db_out.to_excel(out_fn)
print('Mission complete!')

