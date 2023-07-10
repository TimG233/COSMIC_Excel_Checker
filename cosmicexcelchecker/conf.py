# CONFIG FILE TO CUSTOMIZE THE MODULE

# In COSMIC Requirement Spreadsheet, sheet name for demonstrating all CFP points
# default to ['功能点拆分表', 'COSMIC软件评估标准模板']
CFP_SHEET_NAMES = ['功能点拆分表', 'COSMIC软件评估标准模板']

# In NONCOSMIC Requirement Spreadsheet, Sheet name for illustration of workload
NONCFP_SHEET_NAMES = '非COSMIC评估工作量填写说明'

# In COSMIC Requirement Spreadsheet, Column Name for the CFP point, default to 'CFP'
CFP_COLUMN_NAME = 'CFP'

# Subprocess Name, default to '子过程描述'
SUB_PROCESS_NAME = '子过程描述'

# Result Summary Skiprows (结果反馈excel里跳过前..行), default to 9 (due to fixed format)
RS_SKIP_ROWS = 9

# Workload and CFP Ratio, default to 0.79
Workload_CFP_Ratio = 0.79

# Result Summary cosmic workload column name, default to 'cosmic送审工作量'
RS_WORKLOAD_NAME = 'cosmic送审工作量'

# Result Summary cosmic total cfp column name, default to 'cosmic送审功能点'
RS_TOTAL_CFP_NAME = 'cosmic送审功能点'

# Result Summary requirement number column name, default to '需求序号'
RS_REQ_NUM = '需求序号'

# Result Summary requirement name column name, default to '实施需求名称'
RS_REQ_NAME = '实施需求名称'

# Result Summary qualified cosmic column name, default to '是否适用cosmic'
RS_QLF_COSMIC = '是否适用cosmic'

# Single requirement (cosmic) name column name, default to 'OPEX-需求名称'
# This is NOT an EXACT string. This will be checked by startswith() for compatibility
SR_COSMIC_REQ_NAME = 'OPEX-需求名称'

# Single requirement (noncosmic) name column name, default to '需求名称'
SR_NONCOSMIC_REQ_NAME = '需求名称'

# Coefficient sheet name, default to '系数表'
COEFFICIENT_SHEET_NAME = '系数表'

# Coefficient sheet name data column, default to '数值'
COEFFICIENT_SHEET_DATA_COL_NAME = '数值'

# Requirement folder subfolder name, default to 'COSMIC评估发起'
SR_SUBFOLDER_NAME = 'COSMIC评估发起'

# Single Requirement Excel (cosmic) filename (prefix)
SR_COSMIC_FILE_PREFIX = '附件5'

# Single Requirement Excel (non-cosmic) filename (prefix)
SR_NONCOSMIC_FILE_PREFIX = '附件4'

# Single Requirement Excel (non-cosmic) req_num col name, default to '需求序号'
SR_NONCOSMIC_REQ_NUM = '需求序号'

# Single Requirement Excel (non-cosmic) project col name, default to '项目名称'
SR_NONCOSMIC_PROJECT_NAME = '项目名称'