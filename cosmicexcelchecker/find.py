# This file is used to find possible files by requirement

from ._baseclass import UnionExcels

import glob

class FindExcels(UnionExcels):
    '''
    Concrete class for finding all possible Excels under certain paths
    '''

    @staticmethod
    def path_format(path: str) -> str:
        # improve format of path
        return path.replace('\\', '/').strip()

    @staticmethod
    def find_excels(path: str) -> list[str, None]:
        path = FindExcels.path_format(path=path)

        if not path.endswith('/**/*'):
            path += '/**/*'  # to iterate all files

        # using glob for recursive searching
        excels : list[str, None] = []
        for filename in glob.iglob(pathname=path, recursive=True):
            if filename.endswith(('.xlsx', '.xls', '.csv')):
                excels.append(filename.replace('\\', '/'))

        return excels