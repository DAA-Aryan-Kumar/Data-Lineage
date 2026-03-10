import numpy as np
import pandas as pd
import os

input_file_path = os.path.join(os.getcwd(), 'Order Fulfillment Dashboard_Upstream.xlsx')
output_file_path = os.path.join(os.getcwd(), 'formatted_lineage_data.csv')

def format_lineage(lineage_file_path, output_file_path):
    lineage_file = pd.ExcelFile(lineage_file_path)
    lineage_df = lineage_file.parse('Order Fulfillment Dashboard_Ups')
    df = lineage_df.copy()