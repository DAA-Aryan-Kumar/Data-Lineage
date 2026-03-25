import pandas as pd
import os
import argparse
import sys
import logging

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# --- CONFIGURATION: TRIVIAL RULES ---
# You can freely toggle these rules or modify the lists to control what gets expanded in the summary view.
TRIVIAL_RULES = {
    # a. Stop expanding if namespace matching database.schema.name, database.schema, or database
    'enable_namespace_blocking': True,
    'blocked_namespaces': [
        'PROD_DATALAKE.LOGS',        # Example: Blocks everything in LOGS schema
        'TEMP_DB'                    # Example: Blocks entire TEMP_DB
    ],
    
    # b. Stop expanding list of tables if level >= X
    # Put "ALL" in the list for a blanket stop on all tables at/above that level.
    'enable_high_level_tables': True,
    'high_level_tables_limit': 8,
    'high_level_tables': [
        'PROD_EDW.ENT.COMMON_DIM'
    ],
    
    # c. Stop expanding list of views if level >= X
    # Put "ALL" in the list for a blanket stop on all views at/above that level.
    'enable_high_level_views': True,
    'high_level_views_limit': 8,
    'high_level_views': [
        'PROD_EDW.VIEWS.COMMON_VW'
    ],
    
    # Apply these trivial pruning rules to the Detailed Lineage output as well
    'apply_rules_to_detailed': False
}

# --- CONFIGURATION: EXCEL FORMATTING ---
# Control the look and feel of the final exported Excel layout here.
EXCEL_FORMAT = {
    'header_fill_color': 'DDEBF7', # Light blue background for headers (matches typical professional layouts)
    'header_font_color': '000000', # Black text for headers
    'apply_borders': True,
    'border_color': '000000',
    'border_style': 'thin',        # Options: 'thin', 'medium', 'thick'
    'hyperlink_duplicates': True,  # Converts unexpanded duplicate tables into clickable links pointing to their expanded location
    'hyperlink_color': '0000FF'    # Blue color for hyperlinks
}

class DataLineageBuilder:
    def __init__(self, input_file_path: str, output_file_path: str = None):
        if not input_file_path:
            raise ValueError("Input file path must be provided.")
        
        self.input_file_path = os.path.abspath(input_file_path)
        if output_file_path:
            self.output_file_path = os.path.abspath(output_file_path)
        else:
            base_name = self.input_file_path.rsplit(".", 1)[0]
            self.output_file_path = f"{base_name}__Lineage.xlsx"
            
        self.raw_data = None
        self.preprocessed_data = None
        
        self.object_details = None
        
        self.queries_all = None
        self.root_objects_with_upstream = None
        self.root_objects_manual = None
        
        self.lineage = None

    def load_data(self):
        """Loads data from CSV or Excel based on the file extension."""
        if not os.path.exists(self.input_file_path):
            raise FileNotFoundError(f"Input file not found at {self.input_file_path}.")
            
        logging.info(f"Loading data from {self.input_file_path}")
        if self.input_file_path.endswith('.xlsx'):
            raw_input_file = pd.ExcelFile(self.input_file_path)
            self.raw_data = raw_input_file.parse(raw_input_file.sheet_names[0])
        elif self.input_file_path.endswith('.csv'):
            self.raw_data = pd.read_csv(self.input_file_path)
        else:
            raise ValueError("Unsupported file format. Please provide a .csv or .xlsx file.")

    def preprocess_data(self):
        """Preprocesses the raw data to extract consolidated names and structured upstream dependencies."""
        logging.info("Preprocessing data...")
        req_cols = ['Name', 'Database', 'Schema', 'Type', 'Connector', 'Lineage Depth', 'Immediate upstream']
        missing_cols = [col for col in req_cols if col not in self.raw_data.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns in input data: {missing_cols}")

        df = self.raw_data.copy()
        
        df['Consolidated Name'] = df.apply(
            lambda row: f"{row['Database']}.{row['Schema']}.{row['Name']}" if row['Connector'] != 'powerbi' else row['Name'], 
            axis=1
        )
        
        self.object_details = df.drop_duplicates(subset=['Consolidated Name']).set_index('Consolidated Name')
        
        if df['Immediate upstream'].notna().any():
            df['Immediate upstream'] = (
                df['Immediate upstream']
                .dropna()
                .str.split(',')
                .apply(lambda lst: [
                    '.'.join(
                        item.strip().lstrip('(').rstrip(')').split('/')[3:6]
                    )
                    for item in lst
                ] if isinstance(lst, list) else lst)
            )
            df.loc[df['Immediate upstream'].isna(), 'Immediate upstream'] = None

        self.preprocessed_data = df
        
        self.queries_all = df[['Consolidated Name', 'Immediate upstream']].copy()
        self.queries_all.index = pd.RangeIndex(1, len(self.queries_all) + 1)
        
        min_depth = df['Lineage Depth'].min()
        root_objects_df = df[df['Lineage Depth'] == min_depth]
        
        self.root_objects_with_upstream = root_objects_df[root_objects_df['Immediate upstream'].notnull()][['Consolidated Name', 'Immediate upstream']].copy()
        self.root_objects_with_upstream.index = pd.RangeIndex(1, len(self.root_objects_with_upstream) + 1)
        
        self.root_objects_manual = root_objects_df[root_objects_df['Immediate upstream'].isnull()][['Consolidated Name', 'Immediate upstream']].copy()
        self.root_objects_manual.index = pd.RangeIndex(1, len(self.root_objects_manual) + 1)

    def _evaluate_trivial_rules(self, name: str, level: int, rules: dict) -> bool:
        if name not in self.object_details.index:
            return False
        
        details = self.object_details.loc[name]
        
        if rules.get('enable_namespace_blocking'):
            blocked = rules.get('blocked_namespaces', [])
            if any(name.startswith(ns) for ns in blocked):
                return True
                
        obj_type = str(details.get('Type', '')).upper()
        
        if 'TABLE' in obj_type and rules.get('enable_high_level_tables'):
            if level >= rules.get('high_level_tables_limit', 99):
                targets = rules.get('high_level_tables', [])
                if "ALL" in (t.upper() for t in targets) or name in targets:
                    return True
                    
        if 'VIEW' in obj_type and rules.get('enable_high_level_views'):
            if level >= rules.get('high_level_views_limit', 99):
                targets = rules.get('high_level_views', [])
                if "ALL" in (t.upper() for t in targets) or name in targets:
                    return True
                    
        return False

    def build_lineage(self, summarize=False, trivial_rules=None):
        """
        Iteratively traces dependencies backwards to build the pure name lineage hierarchy.
        If summarize=True, expands each node only once, at its lowest depth, and respects trivial rules.
        """
        logging.info(f"Building lineage trace... (Summarize Mode: {summarize})")
        self.lineage = self.root_objects_with_upstream[['Consolidated Name']].copy()
        self.lineage.rename(columns={'Consolidated Name': 'L0'}, inplace=True)
        
        current_level = 1
        global_expanded = set()
        
        # Remove premature L0 seeding from here
        
        while True:
            prev_col = f'L{current_level-1}'
            curr_col = f'L{current_level}'
            
            self.lineage = pd.merge(
                self.lineage,
                self.queries_all[['Immediate upstream', 'Consolidated Name']],
                left_on=prev_col,
                right_on='Consolidated Name',
                how='left'
            )
            self.lineage.drop(columns=['Consolidated Name'], inplace=True)
            self.lineage.rename(columns={'Immediate upstream': curr_col}, inplace=True)
            
            if summarize:
                # 1. Trivial Rules Branch Pruning
                if trivial_rules:
                    skip_mask = self.lineage[prev_col].apply(lambda x: self._evaluate_trivial_rules(x, current_level-1, trivial_rules))
                    self.lineage.loc[skip_mask, curr_col] = None
                
                # 2. Expand-Once Check
                duplicates_mask = self.lineage[prev_col].duplicated(keep='first')
                self.lineage.loc[duplicates_mask, curr_col] = None
                
                already_expanded_mask = self.lineage[prev_col].isin(global_expanded) & (~duplicates_mask)
                self.lineage.loc[already_expanded_mask, curr_col] = None
                
                newly_expanded = self.lineage.loc[self.lineage[curr_col].notna(), prev_col].unique()
                global_expanded.update(newly_expanded)
            
            # Remove self-references
            def filter_circular(r):
                if not isinstance(r[curr_col], list):
                    return r[curr_col]
                seen = {r[f'L{i}'] for i in range(current_level-1, 0, -1) if f'L{i}' in r}
                return [v for v in r[curr_col] if v not in seen]

            self.lineage[curr_col] = self.lineage.apply(filter_circular, axis=1)
            
            self.lineage = self.lineage.explode(curr_col)
            self.lineage.index = pd.RangeIndex(1, len(self.lineage) + 1)
            
            if not self.lineage[curr_col].notna().any():
                self.lineage.drop(columns=[curr_col], inplace=True)
                logging.info(f"No more upstream dependencies found at Level {current_level}. Stopping iteration.")
                break
            
            logging.info(f"Completed pure name Level {current_level}. Total records so far: {len(self.lineage)}")
            current_level += 1
            
    def apply_metadata(self, metadata_cols: list = None):
        """Joins metadata fields iteratively per lineage level into the final output dataframe."""
        if self.lineage is None:
            raise ValueError("Lineage has not been built yet. Call build_lineage() first.")
            
        if metadata_cols is None:
            metadata_cols = ['Type', 'Database', 'Schema', 'Name']
            
        logging.info(f"Applying metadata columns: {metadata_cols}")
        level_cols = [col for col in self.lineage.columns if col.startswith('L')]
        
        final_df = pd.DataFrame(index=self.lineage.index)
        
        for level_col in level_cols:
            final_df[level_col] = self.lineage[level_col]
            for md_col in metadata_cols:
                target_col = f"{md_col}_{level_col}"
                final_df[target_col] = final_df[level_col].map(self.object_details[md_col])
                
        self.lineage = final_df

def reformat_for_client(df):
    """Reorders columns and creates a MultiIndex header to match the client's Excel layout."""
    level_cols = [col for col in df.columns if col.startswith('L') and '_' not in col]
    max_level = len(level_cols) - 1
    
    # We only want L0, then Database_L1, Schema_L1, L1, Type_L1, etc.
    col_order = ['L0']
    for i in range(1, max_level + 1):
        col_order.extend([f'Database_L{i}', f'Schema_L{i}', f'L{i}', f'Type_L{i}'])
        
    # keep only columns that exist
    col_order = [c for c in col_order if c in df.columns]
    df_reordered = df[col_order].copy()
    
    # Build MultiIndex header tuples
    tuples = [('Query', '')]
    for i in range(1, max_level + 1):
        if f'L{i}' in df.columns:
            group_name = 'Sources (L1)' if i == 1 else f'Underlying Sources (L{i})'
            tuples.extend([
                (group_name, 'db'),
                (group_name, 'schema'),
                (group_name, 'object'),
                (group_name, 'type')
            ])
            
    df_reordered.columns = pd.MultiIndex.from_tuples(tuples)
    return df_reordered

def apply_excel_formatting(file_path):
    """Uses openpyxl to apply colors, borders, and hyperlink duplicates."""
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    
    logging.info("Applying openpyxl formatting rules...")
    wb = openpyxl.load_workbook(file_path)
    
    fill = PatternFill(start_color=EXCEL_FORMAT['header_fill_color'], end_color=EXCEL_FORMAT['header_fill_color'], fill_type="solid")
    font = Font(color=EXCEL_FORMAT['header_font_color'], bold=True)
    align_center = Alignment(horizontal='center', vertical='center')
    
    border = None
    if EXCEL_FORMAT['apply_borders']:
        side = Side(border_style=EXCEL_FORMAT['border_style'], color=EXCEL_FORMAT['border_color'])
        border = Border(top=side, bottom=side, left=side, right=side)
        
    hl_font = Font(color=EXCEL_FORMAT['hyperlink_color'], underline="single")
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Pandas MultiIndex requires index=True, which adds an unwanted row index column. Remove it.
        ws.delete_cols(1)
        
        # 1. Format Headers (Rows 1 and 2 because of MultiIndex)
        for row in range(1, 3):
            for cell in ws[row]:
                cell.fill = fill
                cell.font = font
                cell.alignment = align_center
                if border:
                    cell.border = border
                    
        # Apply borders to all data
        if border:
            for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    if cell.value is not None:
                        cell.border = border
                        
        # 2. Hyperlink Duplicates
        if EXCEL_FORMAT['hyperlink_duplicates']:
            first_seen = {}
            # L0 is col 1. L1 object is col 4. L2 object is col 8...
            object_cols = [1] + [4 + (i-1)*4 for i in range(1, 50)]
            
            for row in range(3, ws.max_row + 1):
                for col_idx in object_cols:
                    if col_idx > ws.max_column:
                        break
                        
                    cell = ws.cell(row=row, column=col_idx)
                    val = cell.value
                    if val is not None and str(val).strip() != "":
                        if col_idx == 1:
                            unique_name = str(val).strip()
                        else:
                            db = ws.cell(row=row, column=col_idx-2).value or ""
                            schema = ws.cell(row=row, column=col_idx-1).value or ""
                            unique_name = f"{db}.{schema}.{val}"
                            
                            # Check first_seen
                        if unique_name not in first_seen:
                            first_seen[unique_name] = cell.coordinate
                        else:
                            # It's a duplicate. Map it to the target coordinate
                            target_coord = first_seen[unique_name]
                            cell.value = f'=HYPERLINK("#\'{sheet_name}\'!{target_coord}", "{val}")'
                            cell.font = hl_font
                            cell.style = "Hyperlink"
                            
    wb.save(file_path)

def export_all(detailed_df, summary_df, output_path, separate=False, do_format=True):
    """Exports the lineages and conditionally styles them."""
    if detailed_df is None:
        raise ValueError("Lineages have not been built properly.")
        
    # Reformat columns if formatting is enabled
    if do_format:
        detailed_df = reformat_for_client(detailed_df)
        summary_df = reformat_for_client(summary_df)
        
    try:
        # If we have a MultiIndex (do_format=True), pandas requires index=True. We delete it later in openpyxl.
        write_idx = True if do_format else False
        
        if separate:
            # Export Detailed
            logging.info(f"Saving detailed lineage to {output_path}")
            detailed_df.to_excel(output_path, index=write_idx)
            if do_format: apply_excel_formatting(output_path)
            
            # Export Summary
            sum_path = output_path.replace('.xlsx', '_Summary.xlsx')
            logging.info(f"Saving summarized lineage to {sum_path}")
            summary_df.to_excel(sum_path, index=write_idx)
            if do_format: apply_excel_formatting(sum_path)
        else:
            # Single file, multiple sheets
            logging.info(f"Saving lineage to Excel file (multi-sheet): {output_path}")
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                detailed_df.to_excel(writer, sheet_name='Detailed_Lineage', index=write_idx)
                summary_df.to_excel(writer, sheet_name='Summary_Lineage', index=write_idx)
            if do_format: apply_excel_formatting(output_path)
                
        logging.info("Lineage saved successfully.")
    except Exception as e:
        logging.error(f"Error occurred while saving Excel file: {e}")

def main():
    parser = argparse.ArgumentParser(description="Recursively expand data lineage from Atlan impact reports.")
    parser.add_argument("-i", "--input", type=str, help="Path to the Atlan impact report (.csv or .xlsx)")
    parser.add_argument("-o", "--output", type=str, help="Path to the output Excel file (optional)")
    parser.add_argument("-s", "--separate", action="store_true", help="Export summary lineage to a separate file")
    parser.add_argument("--unformatted", action="store_true", help="Export raw dataframe without Excel styling or MultiIndex layouts")
    
    args = parser.parse_args()
    
    input_file = args.input
    output_file = args.output
    
    if not input_file:
        print("Please enter the path to the Atlan impact report (.csv or .xlsx):")
        try:
            input_file = input("Path: ").strip()
            if input_file.startswith('"') and input_file.endswith('"'):
                input_file = input_file[1:-1]
        except EOFError:
            print("No input provided. Exiting.")
            sys.exit(1)
            
    if not input_file:
        print("No input provided. Exiting.")
        sys.exit(1)

    try:
        builder = DataLineageBuilder(input_file, output_file)
        builder.load_data()
        builder.preprocess_data()
        
        columns_to_include = ['Type', 'Database', 'Schema', 'Name']
        
        # 1. Build Detailed View
        detailed_rules = TRIVIAL_RULES if TRIVIAL_RULES.get('apply_rules_to_detailed', False) else None
        builder.build_lineage(summarize=False, trivial_rules=detailed_rules)
        builder.apply_metadata(metadata_cols=columns_to_include)
        detailed_df = builder.lineage.copy()
        
        # 2. Build Summarized View
        builder.build_lineage(summarize=True, trivial_rules=TRIVIAL_RULES)
        builder.apply_metadata(metadata_cols=columns_to_include)
        summary_df = builder.lineage.copy()
        
        # 3. Export
        do_format = not args.unformatted
        export_all(detailed_df, summary_df, builder.output_file_path, args.separate, do_format)
        
    except Exception as e:
        logging.error(f"Failed to process lineage: {e}")

if __name__ == "__main__":
    main()