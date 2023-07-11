# Core COSMIC File
import numpy
import pandas

from _baseclass import PdExcel
from typing import Union, Dict, List
from errors import CosmicExcelCheckerException ,IncorrectFileTypeException, RepeatedREQNumException, \
    SheetNotFoundException, UnknownREQNumException
from tabulate import tabulate
from conf import CFP_SHEET_NAMES ,CFP_COLUMN_NAME, SUB_PROCESS_NAME, RS_SKIP_ROWS, RS_TOTAL_CFP_NAME, \
    RS_WORKLOAD_NAME, RS_REQ_NUM, RS_REQ_NAME, SR_COSMIC_REQ_NAME,  SR_NONCOSMIC_REQ_NAME, SR_SUBFOLDER_NAME, \
    SR_COSMIC_FILE_PREFIX, SR_NONCOSMIC_FILE_PREFIX, RS_QLF_COSMIC, COEFFICIENT_SHEET_NAME, \
    COEFFICIENT_SHEET_DATA_COL_NAME, NONCFP_SHEET_NAMES, SR_NONCOSMIC_PROJECT_NAME, SR_NONCOSMIC_REQ_NUM, \
    SR_AC_REPORT_NUM, SR_AC_FINAL_NUM, SR_FINAL_CONFIRMATION, SR_AC_REQ_NUM, SR_AC_REQ_NAME, SR_AC_FINAL_NUM_LIMIT

from find import FindExcels

import pandas as pd
import numpy as np
import time
import re
import math

class CosmicReqExcel(PdExcel):
    '''
    Implementation of abstract class PdExcel
    Read a single Excel/CSV about Cosmic Requirement
    '''

    def __init__(self, path: str):
        '''
        data_frames is the pd.DataFrame converted from Spreadsheet
        log holds temporary error/warning for later usage (e.g print to terminal)

        :param path: path to single requirement file
        '''
        self.path : str = path
        self.data_frames: Union[Dict[str, pd.DataFrame], None] = None
        self.log : Union[List[str], str, None] = None

    def load_excel(self):
        file_ext = self.path[self.path.rindex('.'):]

        # it can be simplified to a dict[file_ext:engine] but with less readability
        if file_ext in ('.xlsx', '.xls'):
            self.data_frames = pd.read_excel(self.path, sheet_name=None)
        else:
            raise IncorrectFileTypeException(f"{self.path} is not a valid relative file path for an Excel file")

    def load_csv(self):
        file_ext = self.path[self.path.rindex('.')]

        if file_ext == '.csv':
            self.data_frames = pd.read_csv(self.path)
        else:
            raise IncorrectFileTypeException(f"{self.path} is not a valid relative file path for a csv file")

    def print_df(self):
        '''
        Try to print the converted pd.Dataframe to the terminal

        :return: None
        '''

        for df in self.data_frames.values() if isinstance(self.data_frames, dict) else self.data_frames:

            print(tabulate(df, headers='keys', tablefmt='psql'))

    def get_req_name(self) -> Union[str, None]:
        '''
        Get the requirement name of a single req file

        :return: name string or None if not no data found
        '''

        cfp_df : Union[pd.DataFrame, None] = None
        for sheet_name in CFP_SHEET_NAMES:  # iterate through
            cfp_df = self.data_frames.get(sheet_name, None) if isinstance(self.data_frames, dict) else None

            if cfp_df is not None:
                break

        if cfp_df is None:  # noqa
            return None

        # get column name by iterating through .columns
        for col_name in cfp_df.columns:
            if col_name.startswith(SR_COSMIC_REQ_NAME):
                return cfp_df.iloc[0, cfp_df.columns.get_loc(col_name)]

        return None


    def get_CFP_total(self) -> Union[float, None]:
        '''
        get total CFP pts under CFP column
        it will convert all possible numeric values to dtype float (or int) and leave others as NaN

        :return: total CFP pts
        '''
        data : Union[pd.DataFrame, None] = None
        for sheet_name in CFP_SHEET_NAMES:  # iterate through
            data = self.data_frames.get(sheet_name, None) if isinstance(self.data_frames, dict) else None

            if data is not None:
                break

        if data is None or CFP_COLUMN_NAME not in data.columns or SUB_PROCESS_NAME not in data.columns:
            return None

        cfp_s : pd.Series = pd.to_numeric(data.loc[:, CFP_COLUMN_NAME], errors='coerce')  # convert to numeric
        cfp_df : pd.DataFrame = data.loc[:, [SUB_PROCESS_NAME, CFP_COLUMN_NAME]]

        sub_process_null_count = 0  # only when subprocess is null, not both null
        for index, row in cfp_df.iterrows():
            if pd.isna(row[SUB_PROCESS_NAME]) and not pd.isna(row[CFP_COLUMN_NAME]):
                sub_process_null_count += 1

        return cfp_s.sum() - sub_process_null_count

    def check_CFP_column(self) -> dict:
        '''
        check cfp column with incorrect value
        Correct values: [0, 1]

        :return: a dict-format result
        '''

        for sheet_name in CFP_SHEET_NAMES:  # iterate through
            cfp_df : Union[pd.DataFrame, None] = self.data_frames.get(sheet_name, None) if isinstance(self.data_frames, dict) else None

            if cfp_df is not None:
                break

        if cfp_df is None or CFP_COLUMN_NAME not in cfp_df or SUB_PROCESS_NAME not in cfp_df:  # noqa
            return {'path': self.path, 'match': False, 'CFP': -1, 'note': 'No CFP related column'}

        cfp_df : pd.DataFrame = cfp_df.loc[:, [SUB_PROCESS_NAME, CFP_COLUMN_NAME]]  # noqa
        cfp_total = self.get_CFP_total()  # it will always return a float since CFP column exists

        # record if cfp or sub-process column miss anything
        note = f"Missing data in {'CFP Column' * int(cfp_df.loc[:, CFP_COLUMN_NAME].isna().sum() > 0)} " \
               f"{'Subprocess description' * int(cfp_df.loc[:, SUB_PROCESS_NAME].isna().sum() > 0)}".strip()
        # if cfp_df.loc[:, CFP_COLUMN_NAME].isna().sum() > 0:  # blank space or non-numeric char
        #     note = 'Missing data in CFP Column'
        #
        # if cfp_df.loc[:, SUB_PROCESS_NAME].isna().sum() > 0:
        #     note +=
        if note == 'Missing data in':  # meaning there's no missing
            note = ''

        return {
            "path": self.path,
            "match": True,
            "CFP": cfp_total,
            "note": note
        }

    def check_coefficient_sheet(self) -> Union[bool, None]:
        '''
        check CFP pts is the same with the B1 CFP pts in coefficient sheet if applicable

        :return: match/not match (bool) or None if there's no coefficient_sheet
        '''

        coefficient_sheet: pd.DataFrame = self.data_frames.get(COEFFICIENT_SHEET_NAME, None)

        if coefficient_sheet is None:
            return None

        # print(coefficient_sheet)
        std_cfp_pts = coefficient_sheet.iloc[1, coefficient_sheet.columns.get_loc(COEFFICIENT_SHEET_DATA_COL_NAME)]

        if std_cfp_pts is None or std_cfp_pts == '':
            return None

        try:
            std_cfp_pts = float(std_cfp_pts)

            return std_cfp_pts == self.get_CFP_total()

        except ValueError:
            return None

    def check_final_confirmation(self) -> dict:
        '''
        check final confirmation worksheet if applicable, by the contents comparing to itself and other sheets

        :return: a result dictionary shows the match info
        '''
        for sheet_name in SR_FINAL_CONFIRMATION:
            fc_sheet : Union[pd.DataFrame, None] = self.data_frames.get(sheet_name, None) if \
                isinstance(self.data_frames, dict) else None

            if fc_sheet is not None:
                break

        if fc_sheet is None:  # noqa
            return {
                "path": self.path,
                "match": False,
                "note": "No related final confirmation worksheet found"
            }

        # ignore first row, set index 1 as column and reset index.
        fc_sheet.columns = fc_sheet.iloc[0]  # noqa
        fc_sheet = fc_sheet.iloc[1:]
        fc_sheet = fc_sheet.reset_index(drop=True)

        # load other sheets for comparison
        coefficient_sheet : pd.DataFrame = self.data_frames.get(COEFFICIENT_SHEET_NAME, None) \
            if isinstance(self.data_frames, dict) else None

        if coefficient_sheet is None:
            return {
                "path": self.path,
                "match": False,
                "note": "No related coeffficient worksheet found"
            }

        for sheet_name in CFP_SHEET_NAMES:  # iterate through
            cfp_df : Union[pd.DataFrame, None] = self.data_frames.get(sheet_name, None) \
                if isinstance(self.data_frames, dict) else None

        if cfp_df is None:  # noqa
            return {
                "path": self.path,
                "match": False,
                "note": "No related Cosmic worksheet found"
            }

        try:
            note : str = ''

            # extract folder req by path
            extract_folder_req = rf"\/(?P<req>[0-9]{{1,4}})\/{SR_SUBFOLDER_NAME}\/{SR_COSMIC_FILE_PREFIX}[0-9A-Za-z,&@#$%.\[\]{{}};'\u4e00-\u9fff：。，（）()’￥……]+\.xls[x]{{0,1}}$"
            match = ""
            for m in re.finditer(pattern=extract_folder_req, string=self.path):
                match = m['req']  # type: str

            if match == '':
                note += 'Path Corrupted (no req num subfolder)\t'

            if not match.isnumeric():
                note += 'Missing Req num (cosmic, fc)'

            elif fc_sheet.iloc[0, fc_sheet.columns.get_loc(SR_AC_REQ_NUM)] != int(match):
                    note += 'Req Num not match (cosmic, fc)\t'

            # check req name
            if fc_sheet.iloc[0, fc_sheet.columns.get_loc(SR_AC_REQ_NAME)] != self.get_req_name():
                note += 'Req Name not match (cosmic, fc)\t'

            # # check report num of days
            # std_num_days = round(coefficient_sheet.iloc[-1, -1])
            # if abs(fc_sheet.iloc[0, fc_sheet.columns.get_loc(SR_AC_REPORT_NUM)] - std_num_days) >= 0.1:  # errors
            #     note += 'report num of days not match (coefficient, fc)\t'
            #
            # # check final report num of days
            # if std_num_days <= SR_AC_FINAL_NUM_LIMIT:
            #     final_num_days = std_num_days
            # else:
            #     final_num_days = SR_AC_FINAL_NUM_LIMIT
            #
            # fnd_inchart = fc_sheet.iloc[0, fc_sheet.columns.get_loc(SR_AC_FINAL_NUM)]
            # if fnd_inchart > SR_AC_FINAL_NUM_LIMIT or abs(fnd_inchart - final_num_days) >= 0.1:  # errors
            #     note += 'final num of days not match (fc)\t'
            #
            # note = note.strip('\t')

            return {
                "path": self.path,
                "match": note == '',
                "note": note
            }

        except KeyError:
            return {
                "path": self.path,
                "match": False,
                "note": "Key Error in worksheet. Make sure they are in standard format"
            }


class NonCosmicReqExcel(PdExcel):
    '''
    Implementation of Abstract class PdExcel
    Mainly used for representing non-cosmic Excel file (single requirement)
    '''

    def __init__(self, path: str):
        '''
                data_frames is the pd.DataFrame converted from Spreadsheet
                log holds temporary error/warning for later usage (e.g print to terminal)

                :param path: path to single requirement file
                '''
        self.path: str = path
        self.data_frames: Union[Dict[str, pd.DataFrame], None] = None
        self.log: Union[List[str], str, None] = None

    def load_excel(self):
        file_ext = self.path[self.path.rindex('.'):]

        # it can be simplified to a dict[file_ext:engine] but with less readability
        if file_ext in ('.xlsx', '.xls'):
            self.data_frames = pd.read_excel(self.path, sheet_name=None)
        else:
            raise IncorrectFileTypeException(f"{self.path} is not a valid relative file path for an Excel file")

    def load_csv(self):
        file_ext = self.path[self.path.rindex('.')]

        if file_ext == '.csv':
            self.data_frames = pd.read_csv(self.path)
        else:
            raise IncorrectFileTypeException(f"{self.path} is not a valid relative file path for a csv file")

    def print_df(self):
        '''
        Try to print the converted pd.Dataframe to the terminal

        :return: None
        '''

        for df in self.data_frames.values() if isinstance(self.data_frames, dict) else self.data_frames:
            print(tabulate(df, headers='keys', tablefmt='psql'))

    def get_req_name(self) -> Union[str, None]:
        '''
        Get requirement name for the Excel

        :return: requirement as a string or None if no related column found
        '''

        workload_df = self.data_frames.get(NONCFP_SHEET_NAMES, None)

        if workload_df is None or SR_NONCOSMIC_REQ_NAME not in workload_df.columns:
            return None

        try:
            return workload_df.iloc[1, workload_df.columns.get_loc(SR_NONCOSMIC_REQ_NAME)]
        except IndexError:  # if not exists
            return None

    def get_project_name(self) -> Union[str, None]:
        '''
        Get project name for the Excel

        :return: requirement as a string or None if no related column found
        '''

        workload_df = self.data_frames.get(NONCFP_SHEET_NAMES, None)

        if workload_df is None or SR_NONCOSMIC_PROJECT_NAME not in workload_df.columns:
            return None

        try:
            return workload_df.iloc[1, workload_df.columns.get_loc(SR_NONCOSMIC_PROJECT_NAME)]
        except IndexError:  # if not exists
            return None

class ResultSummary(PdExcel):
    '''
    Another implementation of Abstract class PdExcel
    Demonstrated for loading and processing data in Result Summary related excels
    '''

    def __init__(self, path: str, folders_path:str, sheet_name: str):
        '''
        data_frames is the pd.DataFrame converted from Spreadsheet
        log holds temporary error/warning for later usage (e.g print to terminal)

        :param path: path of the result summary file
        :param sheet_name: specific worksheet name in result summary to be loaded
        '''
        self.path : str = FindExcels.path_format(path=path)
        self.folders_path : str = FindExcels.path_format(path=folders_path)
        self.data_frames: Union[Dict[str, pd.DataFrame], None] = None
        self.data_frame_specific : Union[pd.DataFrame, None] = None
        self.log : Union[list[str], str, None] = None
        self.sheet_name : str = sheet_name
        self.file_paths : Union[list[str, None], None] = FindExcels.find_excels(path=self.folders_path)

    def load_excel(self):
        file_ext = self.path[self.path.rindex('.'):]

        # it can be simplified to a dict[file_ext:engine] but with less readability
        if file_ext in ('.xlsx', '.xls'):
            self.data_frames = pd.read_excel(self.path, sheet_name=None, skiprows=range(RS_SKIP_ROWS))
        else:
            raise IncorrectFileTypeException(f"{self.path} is not a valid relative file path for an Excel file")

    def load_csv(self):
        file_ext = self.path[self.path.rindex('.')]

        if file_ext == '.csv':
            self.data_frames = pd.read_csv(self.path, skiprows=range(RS_SKIP_ROWS))
        else:
            raise IncorrectFileTypeException(f"{self.path} is not a valid relative file path for a csv file")

    def set_sheet_name(self, sheet_name: str):
        '''
        Setting worksheet name for reading dataframe later

        :param sheet_name: str, name of the worksheet
        :return: None
        '''

        self.sheet_name = sheet_name

        df_specific = self.data_frames.get(self.sheet_name, None)

        if df_specific is None:
            raise SheetNotFoundException(f"Sheet with name {sheet_name} is not found inside the given file")

        self.data_frame_specific = df_specific


    def print_df(self):
        '''
        Try to print the converted pd.Dataframe to the terminal

        :return: None
        '''

        for df in self.data_frames.values() if isinstance(self.data_frames, dict) else self.data_frames:

            print(tabulate(df, headers='keys', tablefmt='psql'))


    def print_df_specific(self):
        '''
        Try to print the specific sheet set by function set_sheet_name
        :return: None
        '''

        if isinstance(self.data_frame_specific, pd.DataFrame):
            print(tabulate(self.data_frame_specific, headers='keys', tablefmt='psql'))

        else:
            raise CosmicExcelCheckerException("Specific worksheet is not loaded. Use `set_sheet_name` to load it")

    def check_ratio(self) -> list[str, None]:
        '''
        check all columns workload and cfp ratio, default to 0.79

        :return: a list contains all non-qualified requirements
        '''

        if self.data_frame_specific is None or RS_WORKLOAD_NAME not in self.data_frame_specific.columns or RS_TOTAL_CFP_NAME not in self.data_frame_specific.columns:
            return list()

        result_summary = self.data_frame_specific.loc[:, [RS_REQ_NUM, RS_REQ_NAME, RS_WORKLOAD_NAME,RS_TOTAL_CFP_NAME]]

        bad_ratio : list[str, None] = []  # record
        for index, row in result_summary.iterrows():
            if math.ceil(int(row[RS_WORKLOAD_NAME]) / 0.79) < int(row[RS_TOTAL_CFP_NAME]):
                bad_ratio.append(
                    f'{RS_REQ_NUM}{row[RS_REQ_NUM]}, {RS_REQ_NAME}: {row[RS_REQ_NAME]}'
                )

        return bad_ratio

    def check_file(self, req_num: int) -> dict:
        '''
        check a single file data, comparing to the result summary xlsx
        check req number, req name, CFP total, CFP Total comparison

        :param: req_num is the requirement number 需求序号
        :return: a dict-format result
        '''

        try:
            if self.data_frame_specific is None:
                raise CosmicExcelCheckerException()
        except CosmicExcelCheckerException:
            print("Specific worksheet is not loaded. Use `set_sheet_name` to load it")
            return dict()

        # get row indices qualified for req_num
        r_indices : list = self.data_frame_specific.index[self.data_frame_specific[RS_REQ_NUM] == req_num].tolist()

        try:
            if len(r_indices) < 0:
                raise UnknownREQNumException()

            elif len(r_indices) > 1:
                raise RepeatedREQNumException()

        except UnknownREQNumException:
            print(f"Sheet does not have a requirement number called {req_num}")
            return dict()

        except RepeatedREQNumException:
            print(f"Sheet has repeated rows for requirement number {req_num}")
            return dict()

        row_index = r_indices[0]

        req_num = self.data_frame_specific.iloc[row_index, self.data_frame_specific.columns.get_loc(RS_REQ_NUM)]

        # requirement folder path based on req_num
        req_folder_path = f'{self.folders_path}/{str(req_num)}/'

        req_folder_path_pattern = rf'{self.folders_path}\/(?:.*\/|){str(req_num)}\/'

        # qualified paths inside self.file_paths
        qualified_paths : list = [path for path in self.file_paths if re.match(pattern=req_folder_path_pattern, string=path)]

        if len(qualified_paths) == 0:  # no subfolder found
            return {"REQ Num": req_num, "path": "Not exist", "match": False, "note": "REQ folder does not exist"}

        qualified_cosmic = self.data_frame_specific.iloc[row_index, self.data_frame_specific.columns.get_loc(RS_QLF_COSMIC)]

        # check sr cosmic
        def check_cosmic(path: str):
            if req_num in (86, 98, 110, 208):
                return {}

            # print(req_num)
            cosmic_excel = CosmicReqExcel(path=path)

            # load excel to class df
            cosmic_excel.load_excel()

            note = ''
            # check req name
            if self.data_frame_specific.iloc[
                row_index, self.data_frame_specific.columns.get_loc(RS_REQ_NAME)] != cosmic_excel.get_req_name():
                note += 'REQ name does not match (cosmic)\t'

            # check total CFP name
            total_cfp: str = str(
                self.data_frame_specific.iloc[row_index, self.data_frame_specific.columns.get_loc(RS_TOTAL_CFP_NAME)])
            if total_cfp.isnumeric():
                if float(total_cfp) != cosmic_excel.get_CFP_total():
                    note += 'Total CFP points do not match (cosmic)\t'
            else:
                note += 'CFP points in Result Summary is not valid (cosmic)\t'

            # Check coefficient sheet
            coefficient_sheet_match = cosmic_excel.check_coefficient_sheet()

            if coefficient_sheet_match is None:
                note += 'No Coefficient Sheet in Excel (cosmic)\t'
            elif coefficient_sheet_match is False:
                note += 'Coefficient Sheet B1 data does not match total standard CFP pts\t'

            # check final confirmation worksheet
            fc_result = cosmic_excel.check_final_confirmation()
            note += fc_result['note']

            note = note.rstrip('\t')

            return {"REQ Num": req_num, "path": req_folder_path, "match": note == "", "note": note}

        # check sr noncosmic
        def check_noncosmic(path: str):
            cosmic_excel = NonCosmicReqExcel(path=path)

            # load excel to class df
            cosmic_excel.load_excel()

            note = ''
            # check req name
            if self.data_frame_specific.iloc[
                row_index, self.data_frame_specific.columns.get_loc(RS_REQ_NAME)] != cosmic_excel.get_req_name():
                note += 'REQ name does not match (noncosmic)\t'

            # make sure cfp total is 0 for non-cosmic file
            total_cfp: str = str(
                self.data_frame_specific.iloc[row_index, self.data_frame_specific.columns.get_loc(RS_TOTAL_CFP_NAME)])
            if float(total_cfp) != 0.0:
                note += 'Total CFP points is not 0 for non-cosmic requirement'

            note = note.rstrip('\t')

            return {"REQ Num": req_num, "path": req_folder_path, "match": note == "", "note": ""}

        # qualified paths parent folder
        qualified_paths_docs = [path[path.rindex('/') + 1:] for path in qualified_paths]

        if qualified_cosmic == '是':
            if len(qualified_paths) == 1:

                if not qualified_paths_docs[0].startswith(SR_COSMIC_FILE_PREFIX):
                    return {"REQ Num": req_num, "path": req_folder_path, "match": False, "note": "Incorrect type of cosmic excel based on requirement"}

                # file matched
                return check_cosmic(path=qualified_paths[0])

            else:
                return {"REQ Num": req_num, "path": req_folder_path, "match": False, "note": "Incorrect number of cosmic excel(s) based on requirement"}

        elif qualified_cosmic == '否':
            if len(qualified_paths) == 1:

                if not qualified_paths_docs[0].startswith(SR_NONCOSMIC_FILE_PREFIX):
                    return {"REQ Num": req_num, "path": req_folder_path, "match": False,
                            "note": "Incorrect type of cosmic excel based on requirement"}

                # file matched
                return check_noncosmic(path=qualified_paths[0])

            else:
                return {"REQ Num": req_num, "path": req_folder_path, "match": False,
                        "note": "Incorrect number of cosmic excel(s) based on requirement"}

        elif qualified_cosmic == '混合型':
            if len(qualified_paths) == 2:
                c_prefix = f"{req_folder_path}{SR_COSMIC_FILE_PREFIX}"
                nc_prefix = f"{req_folder_path}{SR_NONCOSMIC_FILE_PREFIX}"

                if not ((qualified_paths[0].startswith(c_prefix) and qualified_paths[1].startswith(nc_prefix)) ^
                        (qualified_paths[0].startswith(nc_prefix) and qualified_paths[1].startswith(c_prefix))):
                    # Same logic to \neq((a==1&b==2)^(a==2&b==1))
                    return {"REQ Num": req_num, "path": req_folder_path, "match": False,
                            "note": "Incorrect type of cosmic excel based on requirement"}

                if qualified_paths_docs[0].startswith(SR_NONCOSMIC_FILE_PREFIX):
                    c_result: dict = check_cosmic(path=qualified_paths[1])
                    nc_result : dict = check_noncosmic(path=qualified_paths[0])
                else:
                    c_result : dict = check_cosmic(path=qualified_paths[0])
                    nc_result : dict = check_noncosmic(path=qualified_paths[1])

                return {"REQ Num": req_num, "path": req_folder_path, "match": c_result['match'] & nc_result['match'],
                        "note": c_result['note'] + nc_result['note']}

            else:
                return {"REQ Num": req_num, "path": req_folder_path, "match": False,
                        "note": "Incorrect number of cosmic excel(s) based on requirement"}

        else:
            return {"REQ Num": req_num, "path": req_folder_path, "match": False,
                    "note": f"The parameter {qualified_cosmic} is not accepted"}

        # # load the cosmic file (needs to improve later)
        # # from here, error will not instantly return, it will continue to leave in 'note' param
        # if qualified_paths[0].startswith(f"{req_folder_path}{SR_NONCOSMIC_FILE_PREFIX}"):
        #     cosmic_excel = CosmicReqExcel(path=qualified_paths[1])
        # else:
        #     cosmic_excel = CosmicReqExcel(path=qualified_paths[0])
        #
        # # load excel to class df
        # cosmic_excel.load_excel()
        #
        # note = ''
        # # check req name
        # if self.data_frame_specific.iloc[row_index, self.data_frame_specific.columns.get_loc(RS_REQ_NAME)] != cosmic_excel.get_req_name():
        #     note += 'REQ name does not match\t'
        #
        # # check total CFP name
        # total_cfp : str = str(self.data_frame_specific.iloc[row_index, self.data_frame_specific.columns.get_loc(RS_TOTAL_CFP_NAME)])
        # if total_cfp.isnumeric():
        #     if float(total_cfp) != cosmic_excel.get_CFP_total():
        #         note += 'Total CFP points do not match\t'
        # else:
        #     note += 'CFP points in Result Summary is not valid'
        #
        # note = note.rstrip('\t')
        #
        # if note != '':  # has invalid info
        #     return {"REQ Num": req_num, "path": req_folder_path, "match": False,
        #             "note": note}
        #
        # return {"REQ Num": req_num, "path": req_folder_path, "match": True, "note": ""}

    def check_all_files(self) -> dict[str, list[dict, None]]:
        '''
        Check all related files listed in the result summary.
        Call `check_file` function for each single check.
        The time complexity is omega(n^2) for calling n items in the Excel of result summary.

        :return: A list of results in dict-format. Could be empty list if nothing found.
        '''

        if self.data_frame_specific is None:
            raise CosmicExcelCheckerException("Specific worksheet is not loaded. Use `set_sheet_name` to load it")

        start_time = time.time()
        # use iterrows() to iterate each row in DataFrame
        list_results : list[dict, None] = []
        for index, row in self.data_frame_specific.iterrows():
            list_results.append(self.check_file(req_num=row[RS_REQ_NUM]))

        cf_results = {
            "results": list_results,
            "time": round(time.time() - start_time, 5)
        }

        return cf_results




# c = CosmicReqExcel(path='附件5：COSMIC功能点拆分表.xls')
c = CosmicReqExcel(path='test.xlsx')
c.load_excel()
# c.print_df()
# print(c.get_CFP_total(data=c.data_frames.get('COSMIC软件评估标准模板', None)))
print(c.get_CFP_total())
print(c.check_CFP_column())
print(c.data_frames.get('功能点拆分表', None).loc[:, 'OPEX-需求名称\nCAPEX-子系统'])
print("req name", c.get_req_name())

fs = FindExcels.find_excels(path='D:\zgydsjy\软件评估\\171')
print(len(fs))
print(FindExcels.find_excels(path='../cosmicexcelchecker'))

# rs = ResultSummary(path='D:\\zgydsjy\\软件评估\\171\\COSMIC结果反馈_20230630.xls', folders_path='D:\\zgydsjy\\软件评估\\171', sheet_name='171_BASS系统软件升级优化服务项目_2023_第二季度')
rs = ResultSummary(path='D:\\zgydsjy\\软件评估\\171\\COSMIC结果反馈_20230630.xls', folders_path='F:\\171', sheet_name='171_BASS系统软件升级优化服务项目_2023_第二季度')
rs.load_excel()
rs.set_sheet_name('171_BASS系统软件升级优化服务项目_2023_第二季度')
# rs.print_df_specific()
print(rs.check_ratio())
print(rs.file_paths)
print(rs.file_paths)
print(len(rs.file_paths))
# rs.data_frame_specific = None
print(rs.check_file(req_num=438))
# print(rs.check_all_files())
resp = rs.check_all_files()
print(resp['time'])
print(resp)

results = resp['results']
for i in range(len(results)):
    try:
        results[i]['REQ Num']
    except KeyError:
        print('err', i)
        print(results[i])

fml = [fm for fm in resp['results'] if fm != dict() and fm['REQ Num'] >= 228 and not fm['match']]
print(fml)
print(len(fml))

c = CosmicReqExcel(path='D:\\zgydsjy\\软件评估\\171\\454\\COSMIC评估发起\\附件5：COSMIC功能点拆分表.xls')
c.load_excel()
print('c check coe', c.check_coefficient_sheet())

