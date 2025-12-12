import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- 1. The Core Transformation Logic ---
def transform_excel(df_a):
  """
  Transforms a DataFrame from Format A to Format B, matching the BA team's output exactly.
  """
  df_a.columns = [col.strip() for col in df_a.columns]
  df_b = pd.DataFrame()
  
  offer_code = df_a.get('Offer Code', pd.Series(dtype='str')).fillna('')
  tnc_no = df_a.get('T&C no.', pd.Series(dtype='str')).fillna('')
  df_b['ITEM NO'] = offer_code + tnc_no
  df_b['Boots_Filename'] = offer_code
  df_b['Barcode'] = df_a.get('Barcode', '')
  
  offer_text_col_name = next((col for col in df_a.columns if 'Offer Text' in col), None)
  def is_use_twice(row):
      if str(row.get('Use Twice?', '')).lower().strip() == 'use twice': return True
      if str(row.get('Use Twice', '')).lower().strip() == 'use twice': return True
      if str(row.get('Part 2', '')).lower().strip() == 'use twice': return True
      if str(row.get('Part 3', '')).lower().strip() == 'use twice': return True
      if offer_text_col_name and str(row.get(offer_text_col_name, '')).lower().strip() == 'use twice': return True
      return False
  
  def determine_layout_type(row):
      p1 = str(row.get('Part 1', ''))
      if p1.lower() == 'save': return 'L2'
      return '(Default)'
  df_b['Layout_Types'] = df_a.apply(determine_layout_type, axis=1)
  
  date_col_name = next((col for col in df_a.columns if 'Date for Coupons' in col), None)
  df_b['Validity'] = df_a[date_col_name].apply(lambda x: f"Valid {x.replace(' to ', '\nto ')}" if x else '') if date_col_name else ''
  
  def format_point1(row):
      if is_use_twice(row): return 'DOUBLE'
      p1 = row.get('Part 1', '')
      if not p1 or pd.isna(p1): return ''
      val_str = str(p1).strip()
      if val_str.lower().endswith('p'): return val_str
      try:
          num_val = float(val_str.replace('£', ''))
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
              return f"{float(val_p2):g}"
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
      processed_text = processed_text.replace('WHEN YOU SPEND', 'WHEN YOU SPEND\n').replace('WHEN YOU BUY', 'WHEN YOU BUY\n').replace('WHEN YOU SHOP', 'WHEN YOU SHOP\n')
      return processed_text
  df_b['Offers'] = df_a.apply(create_offers_text, axis=1)
  
  df_b['_Descriptor'] = df_a.get('Small Print\nInclusions/Exclusions/Medical Information if needed. Use full stop and commas', '')
  
  def format_conditions_1(text):
      if not isinstance(text, str) or not text.strip(): return ''
      lines = [line.strip() for line in text.split('\n') if line.strip()]
      processed_text = '\n\n'.join(lines)
      processed_text = processed_text.replace('please visit\n\n', 'please visit\n')
      return processed_text

  if 'T&Cs Description' in df_a.columns:
      df_b['Conditions_1'] = df_a['T&Cs Description'].apply(format_conditions_1)
  else:
      df_b['Conditions_1'] = ''
  
  def format_offer_type(val):
      if not val or not str(val).startswith('/'): return ''
      try:
          num = int(str(val).replace('/', ''))
          return f'Offer{num}'
      except (ValueError, TypeError): return ''
  df_b['Offer_types'] = df_a.get('T&C no.', pd.Series(dtype='str')).apply(format_offer_type)
      
  df_b['Conditions_3'] = ''
  df_b['_CodeStyles'] = np.where(df_a.get('Barcode', '') != '', 'wCode', 'woCode')
  
  final_columns = ['ITEM NO', 'Layout_Types', 'Validity', 'Point1', 'Point2', 'Point3', 'LogoName', 'Offers', '_Descriptor', 'Offer_types', 'Conditions_1', 'Conditions_3', '_CodeStyles', 'Barcode', 'Boots_Filename']
  df_b = df_b.reindex(columns=final_columns, fill_value='')

  df_b = df_b.applymap(lambda x: x.strip() if isinstance(x, str) else x)

  return df_b

# --- Function to write to Excel with auto-sizing (requires xlsxwriter) ---
def write_excel_with_autosize(df, buffer):
  """Writes DataFrame to an Excel buffer with auto-sized columns and wrapped text."""
  with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
      df.to_excel(writer, index=False, sheet_name='Sheet1')
      
      workbook = writer.book
      worksheet = writer.sheets['Sheet1']
      wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})

      # --- THIS IS THE SIMPLIFIED AND CORRECTED LOGIC ---
      for i, col in enumerate(df.columns):
          # Find the maximum length of the content in the column,
          # considering the longest line in multi-line cells.
          max_content_len = df[col].astype(str).apply(lambda x: max(len(line) for line in x.split('\n'))).max()
          
          # Find the length of the column header itself
          header_len = len(col)
          
          # Set the column width to be the larger of the two, with a little padding
          column_width = max(max_content_len, header_len) + 2
          worksheet.set_column(i, i, column_width, wrap_format)
          
# --- 2. The Streamlit User Interface ---
st.set_page_config(layout="wide", page_title="Boots Coupons Excel Transformation Agent")
col1, col2 = st.columns([1, 6])
with col1:
st.image("Logo.png", width=200)
with col2:
st.title("Boots Coupons Excel Agent")
st.write("This tool converts/transforms the source excel file to the required format for deployment. Please upload your **source file** below.")
st.divider()
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx", help="Upload the source Excel file to be transformed.")
if uploaded_file is not None:
try:
      st.info(f"Processing `{uploaded_file.name}`...")
      input_df = pd.read_excel(uploaded_file, dtype=str).fillna('')
      if 'Offer Code' in input_df.columns:
        input_df = input_df[input_df['Offer Code'].notna() & (input_df['Offer Code'] != '')].copy()
      output_df = transform_excel(input_df)
      st.success("Transformation Complete!")
      output_buffer = BytesIO()
      write_excel_with_autosize(output_df, output_buffer)
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
# --- Sticky Footer ---
footer_css = """
<style>
.footer {
position: fixed; left: 0; bottom: 0; width: 100%;
background-color: #0E1117; color: grey; text-align: center;
padding: 10px; font-size: 0.7em; border-top: 1px solid #262730;
}
.footer a { color: #FF4B4B; text-decoration: none; }
.footer a:hover { text-decoration: underline; }
</style>
"""
footer_html = """
<div class="footer">
Developed by Prince John | Contact <a href='mailto:prince.john@hogarth.com' target="_blank">prince.john@hogarth.com</a> for any assistance
</div>
"""
st.markdown(footer_css + footer_html, unsafe_allow_html=True)
