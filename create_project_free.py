import os
import shutil
import sys
import openpyxl as xl

new_proj_name = sys.argv[1]
dest_path = os.path.join(sys.argv[2], sys.argv[1])
projnr = sys.argv[3]

DOC_PATH = r'C:\Users\EddieKarlsson\Documents'
TEMPLATE_PATH = r'project_structure\free'
source_path = os.path.join(DOC_PATH, TEMPLATE_PATH)

shutil.copytree(source_path, dest_path)

wb_src_path = os.path.join(DOC_PATH, 'MC Doc\WB', 'latest_wb.xlsx')
wb_dst_path = os.path.join(dest_path, 'doc', new_proj_name + '_wb.xlsx')
shutil.copy(wb_src_path, wb_dst_path)

td_src_path = os.path.join(DOC_PATH, 'MC Doc\TD', 'latest_td.xlsx')
td_dst_path = os.path.join(dest_path, 'doc', new_proj_name + '_td.xlsx')
shutil.copy(td_src_path, td_dst_path)

#TODO - open excels with openpyxl and edit projnr and name
