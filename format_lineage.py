import pandas as pd
# import numpy as np
import os

input_file_path = os.path.join(os.getcwd(), 'New Candidate Funnel Dashboard_Upstream.csv')
output_file_path = os.path.join(os.getcwd(), f'{input_file_path.rsplit(".", 1)[0]}__Lineage.xlsx')
save_output = True

def format_lineage():
    if(not os.path.exists(input_file_path)):
        print(f'Input file not found at {input_file_path}. Please check the file path and try again.')
    elif(input_file_path.endswith('.xlsx')):
        raw_input_file = pd.ExcelFile(input_file_path)
        raw_input_sheet = raw_input_file.parse(raw_input_file.sheet_names[0])
    elif(input_file_path.endswith('.csv')):
        raw_input_file = pd.read_csv(input_file_path)
        raw_input_sheet = raw_input_file
    
    preprocessed_input_sheet = raw_input_sheet[['Name','Database','Schema','Type','Connector','Lineage Depth','Immediate upstream']].copy()
    preprocessed_input_sheet['Consolidated Name'] = preprocessed_input_sheet.apply(lambda row: f"{row['Database']}.{row['Schema']}.{row['Name']}" if row['Connector'] != 'powerbi' else row['Name'], axis=1)
    preprocessed_input_sheet['Immediate upstream'] = (
        preprocessed_input_sheet['Immediate upstream']
        .str.split(',')
        .apply(lambda lst: [
            '.'.join(
                item.lstrip('(')
                    .rstrip(')')
                    .split('/')[3:6]
            )
            for item in lst
        ] if isinstance(lst, list) else lst)
    )  if preprocessed_input_sheet['Immediate upstream'].notna().any() else preprocessed_input_sheet['Immediate upstream']

    queries_all = preprocessed_input_sheet[['Consolidated Name','Type','Immediate upstream']].copy()
    queries_all.index = pd.RangeIndex(1, len(queries_all)+1)

    bi_queries_using_snowflake = preprocessed_input_sheet[(preprocessed_input_sheet['Connector'] == 'powerbi') & (preprocessed_input_sheet['Immediate upstream'].notnull())][['Consolidated Name','Type','Immediate upstream']].copy()
    bi_queries_using_snowflake.index = pd.RangeIndex(1, len(bi_queries_using_snowflake)+1)
    bi_queries_manual = preprocessed_input_sheet[(preprocessed_input_sheet['Connector'] == 'powerbi') & (preprocessed_input_sheet['Immediate upstream'].isnull())][['Consolidated Name','Type','Immediate upstream']].copy()
    bi_queries_manual.index = pd.RangeIndex(1, len(bi_queries_manual)+1)
    snowflake_tables = preprocessed_input_sheet[(preprocessed_input_sheet['Connector'] == 'snowflake')][['Consolidated Name','Type','Immediate upstream']].copy()
    snowflake_tables.index = pd.RangeIndex(1, len(snowflake_tables)+1)

    lineage = bi_queries_using_snowflake[['Consolidated Name']].copy()
    lineage.rename(columns={'Consolidated Name': 'L0'}, inplace=True)
    current_level = 1

    while True:
        lineage = pd.merge(
            lineage,
            queries_all[['Immediate upstream','Type','Consolidated Name']],
            left_on=f'L{current_level-1}',
            right_on='Consolidated Name',
            how='left'
        )
        lineage.drop(columns=['Consolidated Name'], inplace=True)
        curr_col = f'L{current_level}'
        lineage.rename(columns={'Immediate upstream': curr_col, 'Type': f'Type_L{current_level}'}, inplace=True)
        
        # remove any value that is the same as the previous levels so we don’t loop infinitely on self-referencing objects
        lineage[curr_col] = lineage.apply(
            lambda r: [v for v in r[curr_col] if v not in {r[f'L{i}'] for i in range(current_level-1,0,-1)}] if isinstance(r[curr_col], list) else r[curr_col],
            axis=1
        )
        
        lineage = lineage.explode(curr_col)
        lineage.index = pd.RangeIndex(1, len(lineage)+1)

        if lineage[curr_col].notna().any(): lineage[[f'Database_L{current_level}', f'Schema_L{current_level}', f'Object_L{current_level}']] = lineage[curr_col].str.split('.', expand=True)
        else:
            lineage.drop(columns=[curr_col, f'Type_L{current_level}'], inplace=True)
            print(f'No more upstream dependencies found at level {current_level}. Stopping iteration.')
            break
        
        print(f'Completed Level {current_level}. Total records so far: {len(lineage)}')
        current_level += 1

    summarized_lineage = lineage.copy()[[col for col in lineage.columns if col.startswith('L') or col.startswith('Type')]].copy()
    if(save_output):
        try:
            print('Saving lineage to Excel file...')
            summarized_lineage.to_excel(output_file_path, index=False)
            print('Lineage saved successfully.')
        except Exception as e:
            print(f'Error occurred while saving Excel file: {e}')