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
pears_export_path = ROOT_DIR + "/example_inputs"

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
# EARS ??? Program Activity Sessions:
# Only program activities that have either more than one session or one
# session greater than or equal to 20 minutes in length are counted.

# Prep Partnerships data
Part_Data = explode_quarters(Part_Data, fy_22_qtr_bounds)

# Prep PSE Site Activities data
PSE_Data = explode_quarters(PSE_Data, fy_22_qtr_bounds)
PSE_NRE_Data = explode_quarters(PSE_NRE_Export, fy_22_qtr_bounds, date_field='baseline_date')
PSE_NRE_Data = PSE_NRE_Data.loc[PSE_NRE_Data['baseline_date'] >= fy_22_qtr_bounds[0]]

# Calculate DHS report metrics


# Function to calculate the quarterly value of a given field
# df: dataframe of PEARS module data with 'report_quarter' column
# field: column used to calculate the quarterly value
# metric: 'sum' or 'count'
# label: string for the column label of the quarterly value
# goals: boolean for whether the metric should be grouped by 'goals' (default: False)
def quarterly_value(df, field, metric, label, goals=False):
    if goals:
        return df.groupby([
            'report_quarter', 'goal'])[field].agg(metric).reset_index(name=label).rename(columns={'goal': 'Goal'})
    else:
        return df.groupby('report_quarter')[field].agg(metric).reset_index(name=label)


# Class that bundles the input arguments of quarterly_value()
class QuarterlyValueInputs:
    def __init__(self, df, field, metric, label, goals=False):
        self.df = df
        self.field = field
        self.metric = metric
        self.label = label
        self.goals = goals


# # of unique programming sites (direct ed & PSE)

unique_sites = PA_Data[['report_quarter', 'snap_ed_grant_goals', 'site_id']].append(
    PSE_Data[['report_quarter', 'snap_ed_grant_goals', 'site_id']], ignore_index=True).drop_duplicates()
unique_sites = explode_goals(unique_sites)
unique_sites = quarterly_value(df=unique_sites,
                               field='site_id',
                               metric='count',
                               label='# of unique programming sites (direct ed & PSE)',
                               goals=True)
# Remove (direct ed & PSE) from column, add to snap_ed_grant_goals?

unique_coalitions = quarterly_value(Coa_Data[['report_quarter', 'coalition_id']].drop_duplicates(),
                                    'coalition_id',
                                    'count',
                                    '# of unique programming sites (direct ed & PSE)')
unique_coalitions['Goal'] = 'Create community collaborations (reported as # coalitions and # organizational members)'
unique_sites = unique_sites.append(unique_coalitions, ignore_index=True)

# Total Unique Reach
# Create package function for total_unique_reach(), pending FY23 guidance

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
reach = quarterly_value(df=site_reach,
                        field='Site Reach',
                        metric='sum',
                        label='Total Reach',
                        goals=True)

# Coa_reach = Coa_Data[
#     ['report_quarter',
#      'number_of_members']].groupby('report_quarter')['number_of_members'].agg('sum').reset_index(name='Total Reach')
# Reach via Coa_Data['number_of_members'] != reach via Coa_Members_Data['member_id']
Coa_reach = quarterly_value(df=Coa_Members_Data[['report_quarter', 'member_id']],
                            field='member_id',
                            metric='count',
                            label='Total Reach')
Coa_reach['Goal'] = 'Create community collaborations (reported as # coalitions and # organizational members)'
reach = reach.append(Coa_reach, ignore_index=True)

goals_sites_reach = pd.merge(unique_sites, reach, how='outer', on=['Goal', 'report_quarter'])

# DE participants reached

PA_demo_dfs = []
demo_subsets = {'participants_total': 'Total',
                'participants_race_amerind': 'American Indian or Alaska Native',
                'participants_race_asian': 'Asian',
                'participants_race_black': 'Black or African American',
                'participants_race_hawpac': 'Native Hawaiian/Other Pacific Islander',
                'participants_race_white': 'White',
                'participants_ethnicity_hispanic': 'Hispanic/Latinx',
                'participants_ethnicity_non_hispanic': 'Non-Hispanic/Non-Latinx'
                }

for demo_field, demo_label in demo_subsets.items():
    PA_demo_dfs.append(quarterly_value(df=PA_Data, field=demo_field, metric='sum', label=demo_label))

# Create function for merging lists of dfs on report_quarter
PA_demo = reduce(lambda left, right: pd.merge(left, right, how='outer', on='report_quarter'), PA_demo_dfs)


# Function to assign a percent column to a dataframe
# df: dataframe used to calculate the percent column
# num: column label to use as the percent numerator
# denom: column label to use as the percent denominator
# label: label for the resulting percent column
def percent(df, num, denom, label):
    df_copy = df.copy()
    df_copy[label] = 100 * df_copy[num] / df_copy[denom]
    return df_copy


for demo_field in demo_subsets.values():
    if demo_field == 'Total':
        continue
    PA_demo = percent(PA_demo, num=demo_field, denom='Total', label='% ' + demo_field)

PA_demo = PA_demo.round(0).drop(columns=['Total'])
# Pivot demo columns into values of Demographic Group column
PA_demo = PA_demo.set_index('report_quarter').stack().reset_index().rename(columns={'level_1': 'Demographic Group'})
PA_demo_a = PA_demo.loc[~PA_demo['Demographic Group'].str.contains('%')]
PA_demo_b = PA_demo.loc[PA_demo['Demographic Group'].str.contains('%')]
PA_demo_b['Demographic Group'] = PA_demo_b['Demographic Group'].str.replace('% ', '')

PA_demo = pd.merge(PA_demo_a, PA_demo_b, how='left', on=['report_quarter', 'Demographic Group']).rename(
    columns={'0_x': 'Total (YTD)', '0_y': '%'})

# RE-AIM Measures of Success

# Reach

reach_inputs = [
    QuarterlyValueInputs(
        df=PA_sites_reach,
        field='PA_participants_sum',
        metric='sum',
        label='# of unique participants attending direct education'),
    QuarterlyValueInputs(
        df=PA_Sessions_Data,
        field='num_participants',
        metric='sum',
        label='# of educational contacts via direct education'),
    QuarterlyValueInputs(
        df=PA_Sessions_Data,
        field='report_quarter',
        metric='count',
        label='# of lessons attended'),
    QuarterlyValueInputs(
        df=IA_IC_Data,
        field='reach',
        metric='sum',
        label='# of indirect education contacts'),
    QuarterlyValueInputs(
        df=PSE_sites_reach,
        field='PSE_total_reach',
        metric='sum',
        label='PSE total estimated reach')
]

reach_dfs = []

for inputs in reach_inputs:
    reach_dfs.append(quarterly_value(inputs.df, inputs.field, inputs.metric, inputs.label))

RE_AIM_Reach = reduce(lambda left, right: pd.merge(left, right, how='outer', on='report_quarter'), reach_dfs)

# Adoption

Part_orgs_reach = quarterly_value(
    df=Part_Data,
    field='partnership_id',
    metric='count',
    label='# of partnering organizations reached'
)

programming_zips_cols = ['report_quarter', 'site_zip']
programming_zips = pd.concat(
    [Coa_Members_Data.loc[Coa_Members_Data['site_zip'].notnull(), programming_zips_cols],
     IA_IC_Data.loc[IA_IC_Data['site_zip'].notnull(), programming_zips_cols],
     PA_Data[programming_zips_cols],
     Part_Data[programming_zips_cols],
     PSE_Data[programming_zips_cols]
     ], ignore_index=True).drop_duplicates()

adoption_zips = quarterly_value(
    df=programming_zips,
    field='site_zip',
    metric='count',
    label='# of unique zip codes reached'
)

RE_AIM_Adoption = pd.merge(Part_orgs_reach, adoption_zips, how='outer', on='report_quarter')

# Implementation

PSE_changes = pd.merge(PSE_Changes_Export, PSE_Data[['pse_id', 'report_quarter']], how='left', on='pse_id')

PSE_change_sites = quarterly_value(
    df=PSE_changes.drop_duplicates(subset=['report_quarter', 'site_id']),
    field='site_id',
    metric='count',
    label='# of sites implementing a PSE change'
)

PSE_NRE_sites = quarterly_value(
    df=PSE_NRE_Data.drop_duplicates(subset=['report_quarter', 'site_id']),
    field='site_id',
    metric='count',
    label='# of unique sites where an organizational readiness or environmental assessment was conducted'
)

RE_AIM_Implementation = pd.merge(PSE_change_sites, PSE_NRE_sites, how='outer', on='report_quarter')

# Count of unique counties/cities with at least 1 partnership entry

Part_Cities = quarterly_value(
    df=Part_Data.drop_duplicates(subset=['report_quarter', 'site_city']),
    field='site_city',
    metric='count',
    label='# of unique cities with at least 1 partnership entry'
)  # Add this metric to report output

Part_Counties = quarterly_value(
    df=Part_Data.drop_duplicates(subset=['report_quarter', 'site_county']),
    field='site_county',
    metric='count',
    label='# of unique counties with at least 1 partnership entry'
)  # Add this metric to report output

# Count and percent of PSE site activities will have a change adopted
# related to food access, diet quality, or physical activity

PSE_changes_count = quarterly_value(
    df=PSE_changes.drop_duplicates(subset=['report_quarter', 'pse_id']),
    field='pse_id',
    metric='count',
    label='# of PSE sites with a plan to implement PSE changes had at least one change adopted'
)

PSE_count = quarterly_value(
    df=PSE_Data,
    field='pse_id',
    metric='count',
    label='# of PSE sites with a plan to implement PSE changes'
)

PSE_changes_count = pd.merge(PSE_changes_count, PSE_count, how='left', on='report_quarter')

# Add this metric to report output
PSE_changes_count = percent(df=PSE_changes_count,
                            num='# of PSE sites with a plan to implement PSE changes had at least one change adopted',
                            denom='# of PSE sites with a plan to implement PSE changes',
                            label='% of PSE sites with a plan to implement PSE changes had at least one change adopted')

# Final Report

report_dfs = [goals_sites_reach, PA_demo, RE_AIM_Reach, RE_AIM_Adoption, RE_AIM_Implementation]

# Check if previous month is the last in the quarter
prev_month = (pd.to_datetime("today") - pd.DateOffset(months=1)).strftime('%m')
fq_lookup = pd.DataFrame({'fq': [1, 2, 3, 4], 'month': ['12', '03', '06', '09']})
if prev_month in fq_lookup['month']:
    current_fq = fq_lookup.loc[fq_lookup['month'] == prev_month, 'fq'].item()
# If not, report all four quarters
else:
    current_fq = 4


# Function to filter a list of dataframes up to the specified fiscal quarter
# dfs: list of dataframes to be filters
# fq_int: integer value for the fiscal quarter
def filter_fq(dfs, fq):
    filtered_dfs = []
    for df in dfs:
        filtered_dfs.append(
            df.loc[df['report_quarter'] <= fq].rename(columns={'report_quarter': 'Report Quarter (YTD)'}))
    return filtered_dfs


report_dfs = filter_fq(report_dfs, current_fq)

out_path = ROOT_DIR + '/example_outputs'
# Create string from fq int instead
filename = out_path + '/DHS Report FY2022 Q' + str(current_fq) + '.xlsx'

tab_names = ['Unique Sites and Reach by Goal', 'Direct Education Demographics', 'RE-AIM Reach', 'RE-AIM Adoption',
             'RE-AIM Implementation']


# Function to export a list of dataframes as an Excel workbook
# file: string for the name or path of the file
# sheet_names: list of strings for the name of each sheet
# dfs: list of dataframes for the report
def write_dhs_report(file, sheet_names, dfs):
    report_dict = dict(zip(sheet_names, dfs))
    writer = pd.ExcelWriter(file, engine='xlsxwriter')
    # Loop through dict of dataframes
    for sheet_name, df in report_dict.items():
        # Send df to writer
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        # Pull worksheet object
        worksheet = writer.sheets[sheet_name]
        # Loop through all columns
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((
                # Len of the largest item
                series.astype(str).map(len).max(),
                # Len of column name/header
                len(str(series.name))
            )) + 1  # adding a little extra space
            # Set column width
            worksheet.set_column(idx, idx, max_len)
            worksheet.autofilter(0, 0, 0, len(df.columns) - 1)
    writer.close()


write_dhs_report(filename, tab_names, report_dfs)
