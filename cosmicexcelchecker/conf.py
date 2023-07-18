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

# Single Requirement Excel (cosmic) analytics confirmation, default to
SR_FINAL_CONFIRMATION = ['结算评估确认表', '结论评估确认表', '结论认同表']

# Analytics confirmation (cosmic) req_num, default to '需求工单号'
SR_AC_REQ_NUM = '需求工单号'

# Analytics confirmation (cosmic) req_name, default to '需求名称'
SR_AC_REQ_NAME = '需求名称'

# Analytics confirmation (cosmic) report number (days), default to '上报工作量（人天）'
SR_AC_REPORT_NUM = '上报工作量\n（人天）'

# Analytics confirmation (cosmic) final result (days), default to '最终结果（人天）'
SR_AC_FINAL_NUM = '最终结果\n（人天）'

# Analytics confirmation (cosmic) final result (days)'s limitation, default to 3.4
SR_AC_FINAL_NUM_LIMIT = 3.4


# allow manually set config
def set_config(config: dict) -> None:
    '''
    Pass in a config dictionaries with {"conf name": value} to change the default value of the config

    :param config: a dict contains configuration
    :return: None
    '''

    if 'CFP_SHEET_NAMES' in config and isinstance(config['CFP_SHEET_NAMES'], (list, tuple)):
        global CFP_SHEET_NAMES
        CFP_SHEET_NAMES = config['CFP_SHEET_NAMES']

    if 'NONCFP_SHEET_NAMES' in config:
        global NONCFP_SHEET_NAMES
        NONCFP_SHEET_NAMES = config['NONCFP_SHEET_NAMES']

    if 'CFP_COLUMN_NAME' in config:
        global CFP_COLUMN_NAME
        CFP_COLUMN_NAME = config['CFP_COLUMN_NAME']

    if 'SUB_PROCESS_NAME' in config:
        global SUB_PROCESS_NAME
        SUB_PROCESS_NAME = config['SUB_PROCESS_NAME']

    if 'RS_SKIP_ROWS' in config:
        global RS_SKIP_ROWS
        RS_SKIP_ROWS = config['RS_SKIP_ROWS']

    if 'Workload_CFP_Ratio' in config:
        global Workload_CFP_Ratio
        Workload_CFP_Ratio = config['Workload_CFP_Ratio']

    if 'RS_WORKLOAD_NAME' in config:
        global RS_WORKLOAD_NAME
        RS_WORKLOAD_NAME = config['RS_WORKLOAD_NAME']

    if 'RS_TOTAL_CFP_NAME' in config:
        global RS_TOTAL_CFP_NAME
        RS_TOTAL_CFP_NAME = config['RS_TOTAL_CFP_NAME']

    if 'RS_REQ_NUM' in config:
        global RS_REQ_NUM
        RS_REQ_NUM = config['RS_REQ_NUM']

    if 'RS_REQ_NAME' in config:
        global RS_REQ_NAME
        RS_REQ_NAME = config['RS_REQ_NAME']

    if 'RS_QLF_COSMIC' in config:
        global RS_QLF_COSMIC
        RS_QLF_COSMIC = config['RS_QLF_COSMIC']

    if 'SR_COSMIC_REQ_NAME' in config:
        global SR_COSMIC_REQ_NAME
        SR_COSMIC_REQ_NAME = config['SR_COSMIC_REQ_NAME']

    if 'SR_NONCOSMIC_REQ_NAME' in config:
        global SR_NONCOSMIC_REQ_NAME
        SR_NONCOSMIC_REQ_NAME = config['SR_NONCOSMIC_REQ_NAME']

    if 'COEFFICIENT_SHEET_NAME' in config:
        global COEFFICIENT_SHEET_NAME
        COEFFICIENT_SHEET_NAME = config['COEFFICIENT_SHEET_NAME']

    if 'COEFFICIENT_SHEET_DATA_COL_NAME' in config:
        global COEFFICIENT_SHEET_DATA_COL_NAME
        COEFFICIENT_SHEET_DATA_COL_NAME = config['COEFFICIENT_SHEET_DATA_COL_NAME']

    if 'SR_SUBFOLDER_NAME' in config:
        global SR_SUBFOLDER_NAME
        SR_SUBFOLDER_NAME = config['SR_SUBFOLDER_NAME']

    if 'SR_COSMIC_FILE_PREFIX' in config:
        global SR_COSMIC_FILE_PREFIX
        SR_COSMIC_FILE_PREFIX = config['SR_COSMIC_FILE_PREFIX']

    if 'SR_NONCOSMIC_FILE_PREFIX' in config:
        global SR_NONCOSMIC_FILE_PREFIX
        SR_NONCOSMIC_FILE_PREFIX = config['SR_NONCOSMIC_FILE_PREFIX']

    if 'SR_NONCOSMIC_REQ_NUM' in config:
        global SR_NONCOSMIC_REQ_NUM
        SR_NONCOSMIC_REQ_NUM = config['SR_NONCOSMIC_REQ_NUM']

    if 'SR_NONCOSMIC_PROJECT_NAME' in config:
        global SR_NONCOSMIC_PROJECT_NAME
        SR_NONCOSMIC_PROJECT_NAME = config['SR_NONCOSMIC_PROJECT_NAME']

    if 'SR_FINAL_CONFIRMATION' in config and isinstance(config['SR_FINAL_CONFIRMATION'], (list, tuple)):
        global SR_FINAL_CONFIRMATION
        SR_FINAL_CONFIRMATION = config['SR_FINAL_CONFIRMATION']

    if 'SR_AC_REQ_NUM' in config:
        global SR_AC_REQ_NUM
        SR_AC_REQ_NUM = config['SR_AC_REQ_NUM ']

    if 'SR_AC_REQ_NAME' in config:
        global SR_AC_REQ_NAME
        SR_AC_REQ_NAME = config['SR_AC_REQ_NAME']

    if 'SR_AC_REPORT_NUM' in config:
        global SR_AC_REPORT_NUM
        SR_AC_REPORT_NUM = config['SR_AC_REPORT_NUM ']

    if 'SR_AC_FINAL_NUM' in config:
        global SR_AC_FINAL_NUM
        SR_AC_FINAL_NUM = config['SR_AC_FINAL_NUM']

    if 'SR_AC_FINAL_NUM_LIMIT' in config:
        global SR_AC_FINAL_NUM_LIMIT
        SR_AC_FINAL_NUM_LIMIT = config['SR_AC_FINAL_NUM_LIMIT']