import numpy as np
import pandas as pd
import os

input_file_path = os.path.join(os.getcwd(), 'Order Fulfillment Dashboard_Upstream.xlsx')
output_file_path = os.path.join(os.getcwd(), 'formatted_lineage_data.csv')

def format_lineage(lineage_file_path, output_file_path):
    lineage_file = pd.ExcelFile(lineage_file_path)
    lineage_sheet = lineage_file.parse('Order Fulfillment Dashboard_Ups')
    # df = lineage_df.copy()
    df = lineage_sheet[['Name','Type','Connector','Database','Schema','Lineage Depth','Immediate upstream']].copy()
    queries = df[(df['Connector'] == 'powerbi') & (df['Immediate upstream'].notnull())][['Name','Immediate upstream']].copy()
    queries.rename(columns={'Name': 'Query', 'Immediate upstream': 'L1'}, inplace=True)
    queries.index = pd.RangeIndex(1, len(queries)+1)
    queries['L1'] = (
        queries['L1']
            .str.extract(r'\((.*?)\)')[0]
            .str.split('/', expand=True)
            .iloc[:, 3:6]
            .apply(lambda row: '.'.join(row.dropna()), axis=1)
    )
    queries[['Database_L1', 'Schema_L1', 'Object_L1']] = queries['L1'].str.split('.', expand=True)
    queries_merged = queries.copy()
    current_level = 2
    while current_level <= 17:  # Arbitrary limit to prevent infinite loops, setting this value to 13 takes about 1 minute to run on my computer, while setting this to 17 will take atleast 7 minutes
        merge_keys = [f'Database_L{current_level-1}', f'Schema_L{current_level-1}', f'Object_L{current_level-1}']
        queries_merged = pd.merge(
            queries_merged,
            df[['Immediate upstream','Database','Schema','Name']],
            left_on=merge_keys,
            right_on=['Database', 'Schema', 'Name'],
            how='left'
        )
        queries_merged.rename(columns={'Immediate upstream': f'L{current_level}', 'Database': f'Database_L{current_level}', 'Schema': f'Schema_L{current_level}', 'Name': f'Object_L{current_level}'}, inplace=True)
        queries_merged[f'L{current_level}'] = queries_merged[f'L{current_level}'].str.split(',')
        prev_col = f'L{current_level-1}'
        curr_col = f'L{current_level}'

        queries_merged = queries_merged.explode(curr_col)
        queries_merged[f'L{current_level}'] = (
            queries_merged[f'L{current_level}']
                .str.extract(r'\((.*?)\)')[0]
                .str.split('/', expand=True)
                .iloc[:, 3:6]
                .apply(lambda row: '.'.join(row.dropna()), axis=1)
        )

        # queries_merged = queries_merged[(queries_merged[curr_col] != queries_merged[prev_col]) | (queries_merged[curr_col] == '')]
        queries_merged.index = pd.RangeIndex(1, len(queries_merged)+1)
        queries_merged[[f'Database_L{current_level}', f'Schema_L{current_level}', f'Object_L{current_level}']] = queries_merged[f'L{current_level}'].str.split('.', expand=True)
        if queries_merged[f'L{current_level}'].isna().all():
            queries_merged.drop(columns=[f'L{current_level}', f'Database_L{current_level}', f'Schema_L{current_level}', f'Object_L{current_level}'], inplace=True)
            break
        print(f'Completed level {current_level}')
        current_level += 1
    queries_merged_cleaned = queries_merged.copy()[['Query'] + [col for col in queries_merged.columns if col.startswith('L')]]