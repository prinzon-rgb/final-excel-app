import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- 1. The Core Transformation Logic ---
def transform_excel(df_a):
  """
  Transforms a DataFrame from Format A to Format B.
  This function now accepts and returns a DataFrame directly.
  """
  # Clean all column names to prevent errors from extra spaces
  df_a.columns = [col.strip() for col in df_a.columns]
  
  df_b = pd.DataFrame()
  
  # Basic Column Mapping
  offer_code = df_a.get('Offer Code', pd.Series(dtype='str')).fillna('')
  tnc_no = df_a.get('T&C no.', pd.Series(dtype='str')).fillna('')
  df_b['ITEM NO'] = offer_code + tnc_no
  df_b['Boots_Filename'] = offer_code
  df_b['Barcode'] = df_a.get('Barcode', '')
  
  # Helper function to find "Use Twice"
  offer_text_col_name = next((col for col in df_a.columns if 'Offer Text' in col), None)
  def is_use_twice(row):
      if str(row.get('Use Twice?', '')).lower().strip() == 'use twice': return True
      if str(row.get('Use Twice', '')).lower().strip() == 'use twice': return True
      if str(row.get('Part 2', '')).lower().strip() == 'use twice': return True
      if str(row.get('Part 3', '')).lower().strip() == 'use twice': return True
      if offer_text_col_name and str(row.get(offer_text_col_name, '')).lower().strip() == 'use twice': return True
      return False
  
  # Layout_Types
  def determine_layout_type(row):
      p1 = str(row.get('Part 1', ''))
      if p1.lower() == 'save': return 'L2'
      return '(Default)'
  df_b['Layout_Types'] = df_a.apply(determine_layout_type, axis=1)
  
  date_col_name = next((col for col in df_a.columns if 'Date for Coupons' in col), None)
  df_b['Validity'] = df_a[date_col_name].apply(lambda x: f"Valid {x.replace(' to ', ' <br/>to ')}" if x else '') if date_col_name else ''
  
  # Points Formatting
  def format_point1(row):
      if is_use_twice(row): return 'DOUBLE'
      p1 = row.get('Part 1', '')
      if not p1 or pd.isna(p1): return ''
      val_str = str(p1)
      if val_str.lower().endswith('p'): return val_str
      try:
          num_val = float(val_str)
          if num_val.is_integer(): return f"£{int(num_val)}"
          return f"£{num_val:.2f}"
      except (ValueError, TypeError): return val_str.upper()
  df_b['Point1'] = df_a.apply(format_point1, axis=1)
  
  def format_point2(row):
      if is_use_twice(row): return 'POINTS'
      val_p1, val_p2 = str(row.get('Part 1', '')), str(row.get('Part 2', ''))
      if not val_p2: return ''
      if val_p1.lower() == 'save':
          try:
              num_val = float(val_p2)
              if np.isclose(num_val, 0.3333333333333333): return '1/3'
              if 0 < num_val < 1: return f"{int(num_val * 100)}%"
              return f"£{float(val_p2):g}"
          except (ValueError, TypeError): return val_p2.upper()
      return val_p2.upper()
  df_b['Point2'] = df_a.apply(format_point2, axis=1)
  
  df_b['Point3'] = df_a.apply(lambda row: 'USE TWICE' if is_use_twice(row) else '', axis=1)
  df_b['LogoName'] = df_a.get('Logo', pd.Series(dtype='str')).apply(lambda x: f"{x}.pdf" if x and str(x).lower() not in ['n/a', ''] else '')
  
  def create_offers_text(row):
      if is_use_twice(row): return ''
      offer_text = row.get(offer_text_col_name, '') if offer_text_col_name else ''
      if not offer_text: return ''
      processed_text = str(offer_text).replace('\n', ' ')
      processed_text = processed_text.upper().replace('NO7', 'No7')
      processed_text = processed_text.replace('WHEN YOU SPEND', 'WHEN YOU SPEND<br/>').replace('WHEN YOU BUY', 'WHEN YOU BUY<br/>').replace('WHEN YOU SHOP', 'WHEN YOU SHOP<br/>')
      return processed_text
  df_b['Offers'] = df_a.apply(create_offers_text, axis=1)
  
  # Descriptor & T&Cs
  small_print_col_name = next((col for col in df_a.columns if 'Small Print' in col), None)
  df_b['_Descriptor'] = df_a.get(small_print_col_name, '') if small_print_col_name else ''
  
  def format_conditions_1(text):
      if not isinstance(text, str): return ''
      lines = text.split('\n')
      result_parts = []
      i = 0
      while i < len(lines):
          current_line = lines[i]
          if (i + 1 < len(lines)) and ('boots.com' in lines[i+1] or 'www.' in lines[i+1]):
              result_parts.append(current_line + ' ' + lines[i+1].strip())
              i += 2
          else:
              result_parts.append(current_line)
              i += 1
      return '<br/>'.join(result_parts)
  
  if 'T&Cs Description' in df_a.columns:
      df_b['Conditions_1'] = df_a['T&Cs Description'].apply(format_conditions_1)
  else:
      df_b['Conditions_1'] = ''
  
  def format_offer_type(val):
      if not val or not str(val).startswith('/'): return ''
      try:
          num = int(str(val).replace('/', ''))
          return f'Offer{num:02d}'
      except (ValueError, TypeError): return ''
  df_b['Offer_types'] = df_a.get('T&C no.', pd.Series(dtype='str')).apply(format_offer_type)
  
  df_b['Conditions_3'] = ''
  df_b['_CodeStyles'] = np.where(df_a.get('Barcode', '') != '', 'wCode', 'woCode')
  
  final_columns = ['ITEM NO', 'Layout_Types', 'Validity', 'Point1', 'Point2', 'Point3', 'LogoName', 'Offers', '_Descriptor', 'Offer_types', 'Conditions_1', 'Conditions_3', '_CodeStyles', 'Barcode', 'Boots_Filename']
  df_b = df_b.reindex(columns=final_columns, fill_value='')
  
  return df_b

# --- 2. The Streamlit User Interface ---
st.set_page_config(layout="centered", page_title="Excel Transformation Agent")
st.title("📄 Excel Transformation Agent")
st.write(
"This tool automates the transformation of Excel files. Upload your file below to begin."
)
st.divider()

uploaded_file = st.file_uploader(
"Choose an Excel file",
type="xlsx",
help="Upload the source Excel file to be transformed."
)

if uploaded_file is not None:
  try:
        st.info(f"Processing `{uploaded_file.name}`...")
    
        input_df = pd.read_excel(uploaded_file, dtype=str).fillna('')
    
        if 'Offer Code' in input_df.columns:
          input_df = input_df[input_df['Offer Code'].notna() & (input_df['Offer Code'] != '')].copy()

        output_df = transform_excel(input_df)
    
        st.success("Transformation Complete!")

        output_buffer = BytesIO()
        output_df.to_excel(output_buffer, index=False, sheet_name='Sheet1')
        output_buffer.seek(0)
    
        st.download_button(
          label="⬇️ Download Transformed File",
          data=output_buffer,
          file_name=f"{uploaded_file.name.replace('.xlsx', '')}-transformed.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.subheader("Preview of Transformed Data")
        st.dataframe(output_df)
    
  except Exception as e:
        st.error(f"An error occurred: {e}")
        st.warning("Please ensure the uploaded file has a compatible structure.")
