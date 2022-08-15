import os
import pandas as pd
from datetime import datetime
from functools import reduce

# Calculate the path to the root directory of this script
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))

# Define path to directory for reformatted PEARS module exports
# Used output path from pears_nightly_export_reformatting.py
# Otherwise, custom field labels will cause errors
# pears_export_path = r"\path\to\reformatted_pears_data"
# Script demo uses /example_inputs directory
pears_export_path = pd.ExcelFile(ROOT_DIR + "/example_inputs")

# Import Indirect Activity data and Intervention Channels
Indirect_Activities_Export = pd.ExcelFile(pears_export_path + '/' + "Indirect_Activity_Export.xlsx")
IA_Export = pd.read_excel(Indirect_Activities_Export, 'Indirect Activity Data')
# Only report on records for SNAP-Ed
IA_Data = IA_Export.loc[
    (IA_Export['program_area'] == 'SNAP-Ed') & (~IA_Export['title'].str.contains('(?i)TEST', regex=True))]
IA_IC_Export = pd.read_excel(Indirect_Activities_Export, 'Intervention Channels')

# Import Coalitions data and Coalition Members
Coalitions_Export = pd.ExcelFile(pears_export_path + '/' + "Coalition_Export.xlsx")
Coa_Export = pd.read_excel(Coalitions_Export, 'Coalition Data')
# Only report on records for SNAP-Ed
Coa_Data = Coa_Export.loc[
    (Coa_Export['program_area'] == 'SNAP-Ed') & (~Coa_Export['coalition_name'].str.contains('(?i)TEST', regex=True))]
Coa_Members_Export = pd.read_excel(Coalitions_Export, 'Members')

# Import Program Activity data and Sessions
Program_Activities_Export = pd.ExcelFile(pears_export_path + '/' + "Program_Activities_Export.xlsx")
PA_Export = pd.read_excel(Program_Activities_Export, 'Program Activity Data')
# PA is only module to have cross-program_area collaboration
PA_Data = PA_Export.loc[
    (PA_Export['program_areas'].str.contains('SNAP-Ed')) & (~PA_Export['name'].str.contains('(?i)TEST', regex=True))]
PA_Sessions_Export = pd.read_excel(Program_Activities_Export, 'Sessions')

# Import Partnerships data
Partnerships_Export = pd.ExcelFile(pears_export_path + '/' + "Partnership_Export.xlsx")
Part_Export = pd.read_excel(Partnerships_Export, 'Partnership Data')
# Only report on records for SNAP-Ed
Part_Data = Part_Export.loc[(Part_Export['program_area'] == 'SNAP-Ed') & (
    ~Part_Export['partnership_name'].str.contains('(?i)TEST', regex=True))]

# Import PSE Site Activity data, Needs, Readiness, Effectiveness, and Changes
PSE_Site_Activities_Export = pd.ExcelFile(pears_export_path + '/' + "PSE_Site_Activity_Export.xlsx")
PSE_Export = pd.read_excel(PSE_Site_Activities_Export, 'PSE Data')
PSE_Data = PSE_Export.loc[~PSE_Export['name'].str.contains('(?i)TEST', regex=True, na=False)]
PSE_Changes_Export = pd.read_excel(PSE_Site_Activities_Export, 'Changes')
PSE_NRE_Export = pd.read_excel(PSE_Site_Activities_Export, 'Needs, Readiness, Effectiveness')

# Assign Quarters

fy_22_qtr_bounds = ['10/01/2021', '01/11/2022', '04/11/2022', '07/11/2022', '10/18/2022']


# Function for assigning and exploding records by quarter
# df: dataframe of module export sheet records
# qtr_bounds: list of date strings for the lower/upper bounds of each quarter
# date_field: df column name of the date field to base quarters on (default: 'created')
def explode_quarters(df, qtr_bounds, date_field='created'):
    df[date_field] = pd.to_datetime(df[date_field])  # SettingWithCopyWarning
    df['report_quarter'] = ''  # SettingWithCopyWarning
    dt_list = [datetime.strptime(s, "%m/%d/%Y") for s in qtr_bounds]
    # Upper bound is exclusive
    df.loc[(df[date_field] >= dt_list[0]) & (df[date_field] < dt_list[1]), 'report_quarter'] = '1, 2, 3, 4'
    df.loc[(df[date_field] >= dt_list[1]) & (df[date_field] < dt_list[2]), 'report_quarter'] = '2, 3, 4'
    df.loc[(df[date_field] >= dt_list[2]) & (df[date_field] < dt_list[3]), 'report_quarter'] = '3, 4'
    df.loc[(df[date_field] >= dt_list[3]) & (df[date_field] < dt_list[4]), 'report_quarter'] = '4'
    df['report_quarter'] = df['report_quarter'].str.split(', ').tolist()  # SettingWithCopyWarning
    df = df.explode('report_quarter')
    df['report_quarter'] = pd.to_numeric(df['report_quarter'])
    return df


# function for exploding snap_ed_grant_goals field
# df: dataframe of module export sheet records
def explode_goals(df):
    df['goal'] = df['snap_ed_grant_goals'].str.split(',').tolist()
    df = df.explode('goal')
    return df


# Prep Coalitions data
Coa_Data = explode_quarters(Coa_Data, fy_22_qtr_bounds)
Coa_Members_Data = pd.merge(Coa_Members_Export, Coa_Data[['coalition_id', 'program_area', 'report_quarter']],
                            how='left', on='coalition_id')
Coa_Members_Data = Coa_Members_Data.loc[Coa_Members_Data['program_area'] == 'SNAP-Ed']

# Prep Indirect Activities data
IA_Data = explode_quarters(IA_Data, fy_22_qtr_bounds)
IA_IC_Data = pd.merge(IA_IC_Export, IA_Data[['activity_id', 'program_area', 'report_quarter']], how='left',
                      on='activity_id')
IA_IC_Data = IA_IC_Data.loc[IA_IC_Data['program_area'] == 'SNAP-Ed']
# Use IC created field when export updated

# Prep Program Activities data
PA_Data = explode_quarters(PA_Data, fy_22_qtr_bounds)
PA_Sessions_Data = pd.merge(PA_Sessions_Export, PA_Data[['program_id', 'program_areas']].drop_duplicates(), how='left',
                            on='program_id')
PA_Sessions_Data = PA_Sessions_Data.loc[PA_Sessions_Data['program_areas'].str.contains('SNAP-Ed', na=False)]
PA_Sessions_Data = explode_quarters(PA_Sessions_Data, fy_22_qtr_bounds, date_field='start_date')
PA_Sessions_Data = PA_Sessions_Data.loc[PA_Sessions_Data['report_quarter'] != '']
# EARS – Program Activity Sessions:
# Only program activities that have either more than one session or one
# session greater than or equal to 20 minutes in length are counted.

# Prep Partnerships data
Part_Data = explode_quarters(Part_Data, fy_22_qtr_bounds)

# Prep PSE Site Activities data
PSE_Data = explode_quarters(PSE_Data, fy_22_qtr_bounds)
PSE_NRE_Data = explode_quarters(PSE_NRE_Export, fy_22_qtr_bounds, date_field='baseline_date')
PSE_NRE_Data = PSE_NRE_Data.loc[PSE_NRE_Data['baseline_date'] >= fy_22_qtr_bounds[0]]

# Calculate DHS report metrics

# # of unique programming sites (direct ed & PSE)

unique_sites = PA_Data[['report_quarter', 'snap_ed_grant_goals', 'site_id']].append(
    PSE_Data[['report_quarter', 'snap_ed_grant_goals', 'site_id']], ignore_index=True).drop_duplicates()
unique_sites = explode_goals(unique_sites)
unique_sites = unique_sites.groupby(['report_quarter', 'goal'])['site_id'].count().reset_index(
    name='# of unique programming sites (direct ed & PSE)')
unique_sites = unique_sites.rename(columns={'goal': 'Goal'})
# Remove (direct ed & PSE) from column, add to snap_ed_grant_goals?

unique_coalitions = Coa_Data[['report_quarter', 'coalition_id']].drop_duplicates().groupby('report_quarter')[
    'coalition_id'].count().reset_index(name='# of unique programming sites (direct ed & PSE)')
unique_coalitions['Goal'] = 'Create community collaborations (reported as # coalitions and # organizational members)'
unique_sites = unique_sites.append(unique_coalitions, ignore_index=True)

# Total Unique Reach

PA_sites_reach = PA_Data[['report_quarter', 'snap_ed_grant_goals', 'site_id', 'participants_total']]
PA_sites_reach = explode_goals(PA_sites_reach)
PA_sites_reach = PA_sites_reach.rename(columns={'participants_total': 'PA_participants_total'}).groupby(
    ['report_quarter', 'goal', 'site_id'])['PA_participants_total'].agg('sum').reset_index(name='PA_participants_sum')
PSE_sites_reach = explode_goals(PSE_Data.loc[
                                    PSE_Data['total_reach'].notnull(), ['report_quarter', 'snap_ed_grant_goals',
                                                                        'site_id', 'total_reach']])
PSE_sites_reach = PSE_sites_reach.sort_values(['report_quarter', 'site_id', 'total_reach']).rename(
    columns={'total_reach': 'PSE_total_reach'}).drop_duplicates(subset=['report_quarter', 'site_id'], keep='last')
site_reach = pd.merge(PA_sites_reach, PSE_sites_reach, how='outer', on=['report_quarter', 'goal', 'site_id'])
site_reach['Site Reach'] = site_reach[['PA_participants_sum', 'PSE_total_reach']].max(axis=1)
reach = site_reach.groupby(['report_quarter', 'goal'])['Site Reach'].agg('sum').reset_index(
    name='Total Reach')  # Reach (greater of direct ed or PSE across unique sites)
reach = reach.rename(columns={'goal': 'Goal'})

# Coa_reach = Coa_Data[
#     ['report_quarter',
#      'number_of_members']].groupby('report_quarter')['number_of_members'].agg('sum').reset_index(name='Total Reach')
# Reach via Coa_Data['number_of_members'] != reach via Coa_Members_Data['member_id']
Coa_reach = Coa_Members_Data[['report_quarter', 'member_id']].groupby('report_quarter')[
    'member_id'].count().reset_index(name='Total Reach')
Coa_reach['Goal'] = 'Create community collaborations (reported as # coalitions and # organizational members)'
reach = reach.append(Coa_reach, ignore_index=True)

goals_sites_reach = pd.merge(unique_sites, reach, how='outer', on=['Goal', 'report_quarter'])

# DE participants reached

# Create a function for the copied lines below
# pa_df:
# participants_fields:
# participants_subsets:
PA_total_participants = PA_Data.groupby('report_quarter')['participants_total'].agg('sum').reset_index(name='Total')
PA_amerind = PA_Data.groupby('report_quarter')['participants_race_amerind'].agg('sum').reset_index(
    name='American Indian or Alaska Native')
PA_asian = PA_Data.groupby('report_quarter')['participants_race_asian'].agg('sum').reset_index(name='Asian')
PA_black = PA_Data.groupby('report_quarter')['participants_race_black'].agg('sum').reset_index(
    name='Black or African American')
PA_hawpac = PA_Data.groupby('report_quarter')['participants_race_hawpac'].agg('sum').reset_index(
    name='Native Hawaiian/Other Pacific Islander')
PA_white = PA_Data.groupby('report_quarter')['participants_race_white'].agg('sum').reset_index(name='White')

PA_race = reduce(lambda left, right: pd.merge(left, right, how='outer', on='report_quarter'),
                 [PA_total_participants, PA_amerind, PA_asian, PA_black, PA_hawpac, PA_white])

# Create a function for the copied lines below
PA_race['% American Indian or Alaska Native'] = 100 * PA_race['American Indian or Alaska Native'] / PA_race['Total']
PA_race['% Asian'] = 100 * PA_race['Asian'] / PA_race['Total']
PA_race['% Black or African American'] = 100 * PA_race['Black or African American'] / PA_race['Total']
PA_race['% Native Hawaiian/Other Pacific Islander'] = 100 * PA_race['Native Hawaiian/Other Pacific Islander'] / PA_race[
    'Total']
PA_race['% White'] = 100 * PA_race['White'] / PA_race['Total']

PA_hispanic = PA_Data.groupby('report_quarter')['participants_ethnicity_hispanic'].agg('sum').reset_index(
    name='Hispanic/Latinx')
PA_non_hispanic = PA_Data.groupby('report_quarter')['participants_ethnicity_non_hispanic'].agg('sum').reset_index(
    name='Non-Hispanic/Non-Latinx')
PA_ethnicity = pd.merge(PA_hispanic, PA_non_hispanic, how='outer', on='report_quarter')

PA_demo = pd.merge(PA_race, PA_ethnicity, how='outer', on='report_quarter')
PA_demo['% Hispanic/Latinx'] = 100 * PA_demo['Hispanic/Latinx'] / PA_demo['Total']
PA_demo['% Non-Hispanic/Non-Latinx'] = 100 * PA_demo['Non-Hispanic/Non-Latinx'] / PA_demo['Total']
PA_demo = PA_demo.round(0).drop(columns=['Total'])

PA_demo = PA_demo.set_index('report_quarter').stack().reset_index().rename(columns={'level_1': 'Demographic Group'})

PA_demo_a = PA_demo.loc[~PA_demo['Demographic Group'].str.contains('%')]
PA_demo_b = PA_demo.loc[PA_demo['Demographic Group'].str.contains('%')]
PA_demo_b['Demographic Group'] = PA_demo_b['Demographic Group'].str.replace('% ', '')

PA_demo = pd.merge(PA_demo_a, PA_demo_b, how='left', on=['report_quarter', 'Demographic Group']).rename(
    columns={'0_x': 'Total (YTD)', '0_y': '%'})

# RE-AIM Measures of Success

# Reach

PA_unique_reach = PA_sites_reach.groupby('report_quarter')['PA_participants_sum'].agg('sum').reset_index(
    name='# of unique participants attending direct education')
PA_ed_touches = PA_Sessions_Data.groupby('report_quarter')['num_participants'].agg('sum').reset_index(
    name='# of educational contacts via direct education')
PA_lessons = PA_Sessions_Data.groupby('report_quarter')['session_id'].agg('count').reset_index(
    name='# of lessons attended')
IA_ed_touches = IA_IC_Data.groupby('report_quarter')['reach'].agg('sum').reset_index(
    name='# of indirect education contacts')
# Include Find Food IL map reach from Google Analytics?
# From IA process sheet: "Report statewide entries individually. Report all other IA entries in aggregate." ?
PSE_reach = PSE_sites_reach.groupby('report_quarter')['PSE_total_reach'].agg('sum').reset_index(
    name='PSE total estimated reach')

RE_AIM_Reach = reduce(lambda left, right: pd.merge(left, right, how='outer', on='report_quarter'),
                      [PA_unique_reach, PA_ed_touches, PA_lessons, IA_ed_touches, PSE_reach])

# Adoption

Part_orgs_reach = Part_Data.groupby('report_quarter')['partnership_id'].count().reset_index(
    name='# of partnering organizations reached')

programming_zips = pd.concat(
    [Coa_Members_Data.loc[Coa_Members_Data['site_zip'].notnull(), ['report_quarter', 'site_zip']],
     IA_IC_Data.loc[IA_IC_Data['site_zip'].notnull(), ['report_quarter', 'site_zip']],
     PA_Data[['report_quarter', 'site_zip']],
     Part_Data[['report_quarter', 'site_zip']],
     PSE_Data[['report_quarter', 'site_zip']]], ignore_index=True)

adoption_zips = programming_zips.drop_duplicates().groupby('report_quarter')['site_zip'].count().reset_index(
    name='# of unique zip codes reached')

RE_AIM_Adoption = pd.merge(Part_orgs_reach, adoption_zips, how='outer', on='report_quarter')

# Implementation

PSE_changes = pd.merge(PSE_Changes_Export, PSE_Data[['pse_id', 'report_quarter']], how='left', on='pse_id')
PSE_change_sites = PSE_changes.drop_duplicates(subset=['report_quarter', 'site_id']).groupby('report_quarter')[
    'site_id'].count().reset_index(name='# of sites implementing a PSE change')

PSE_NRE_sites = PSE_NRE_Data.drop_duplicates(subset=['report_quarter', 'site_id']).groupby('report_quarter')[
    'site_id'].count().reset_index(
    name='# of unique sites where an organizational readiness or environmental assessment was conducted')

RE_AIM_Implementation = pd.merge(PSE_change_sites, PSE_NRE_sites, how='outer', on='report_quarter')

# Count of unique counties/cities with at least 1 partnership entry

Part_Cities = Part_Data.drop_duplicates(subset=['report_quarter', 'site_city']).groupby('report_quarter')[
    'site_city'].count().reset_index(name='# of unique cities with at least 1 partnership entry.')
Part_Counties = Part_Data.drop_duplicates(subset=['report_quarter', 'site_county']).groupby('report_quarter')[
    'site_county'].count().reset_index(name='# of unique counties with at least 1 partnership entry.')

# Count and percent of PSE site activities will have a change adopted
# related to food access, diet quality, or physical activity

PSE_changes_count = PSE_changes.drop_duplicates(subset=['report_quarter', 'pse_id']).groupby('report_quarter')[
    'pse_id'].count().reset_index(
    name='# of PSE sites with a plan to implement PSE changes had at least one change adopted.')
PSE_count = PSE_Data.groupby('report_quarter')['pse_id'].count().reset_index(
    name='# of PSE sites with a plan to implement PSE changes.')
PSE_changes_count = pd.merge(PSE_changes_count, PSE_count, how='left', on='report_quarter')
PSE_changes_count['% of PSE sites with a plan to implement PSE changes had at least one change adopted.'] = 100 * \
                                                                                                            PSE_changes_count[
                                                                                                                '# of PSE sites with a plan to implement PSE changes had at least one change adopted.'] / \
                                                                                                            PSE_changes_count[
                                                                                                                '# of PSE sites with a plan to implement PSE changes.']

# Final Report

prev_month = (pd.to_datetime("today") - pd.DateOffset(months=1)).strftime('%m')
fq_lookup = pd.DataFrame({'fq': ['Q1', 'Q2', 'Q3', 'Q4'], 'month': ['12', '03', '06', '09'], 'int': [1, 2, 3, 4]})
fq = fq_lookup.loc[fq_lookup['month'] == prev_month, 'fq'].item()
fq_int = fq_lookup.loc[fq_lookup['month'] == prev_month, 'int'].item()

dfs = [goals_sites_reach, PA_demo, RE_AIM_Reach, RE_AIM_Adoption, RE_AIM_Implementation]
filtered_dfs = []

for df in dfs:
    filtered_dfs.append(
        df.loc[df['report_quarter'] <= fq_int].rename(columns={'report_quarter': 'Report Quarter (YTD)'}))

filename = 'DHS Report FY2022 ' + fq + '.xlsx'

tab_names = ['Unique Sites and Reach by Goal', 'Direct Education Demographics', 'RE-AIM Reach', 'RE-AIM Adoption',
             'RE-AIM Implementation']

dfs_dict = dict(zip(tab_names, filtered_dfs))

writer = pd.ExcelWriter(filename, engine='xlsxwriter')
for sheetname, df in dfs_dict.items():  # loop through `dict` of dataframes
    df.to_excel(writer, sheet_name=sheetname, index=False)  # send df to writer
    worksheet = writer.sheets[sheetname]  # pull worksheet object
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
        )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
        worksheet.autofilter(0, 0, 0, len(df.columns) - 1)

writer.close()
