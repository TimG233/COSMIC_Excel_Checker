# _baseclass to include all abstract class for the module

from abc import ABC, abstractmethod

class PdExcel(ABC):
    '''
    An Abstract class demonstrated the base class related to the excel_files loaded
    All related file with concrete methods will implement this base class
    '''

    @abstractmethod
    def __init__(self, path: str):
        '''
        Instantiation of the class, which is an abstract method in this case
        :param path: RELATIVE path to the Excel file, file extension xlsx is needed
        '''

        pass

    @abstractmethod
    def load_excel(self):
        '''
        Load excel all possible spreedsheets from Excel file
        :return: None
        '''

        pass

    @abstractmethod
    def load_csv(self):
        '''
        Load csv plaintext file
        :return: None
        '''

        pass

class UnionExcels(ABC):
    '''
    An abstract class designed for finding all Excel files under certain directory
    This will be implemented by classes with concrete methods
    '''

    @staticmethod
    @abstractmethod
    def path_format(path: str) -> str:
        '''
        format path by replacing '\' or '\\' to '/'
        :param path:
        :return: formatted path
        '''

        pass

    @staticmethod
    @abstractmethod
    def find_excels(path: str) -> list[str, None]:
        '''
        find all possible Excel files under path (in constructor) and return name of them as a list
        :path: relative path for searching files
        :return: list of all Excel fils name (path) or list of 0 element if no files qualified
        '''

        pass

class AbstractObf(ABC):
    '''
    An abstract class designed for testing any obfuscation (fuzzy)
    This will be implemented by classes with concrete methods
    '''

    @staticmethod
    @abstractmethod
    def compare(string1: str, string2: str) -> int:
        '''
        compare two strings

        :param string1: string to check obfuscation
        :param string2: another string to check obfuscation
        :return: the number of edit distance as integer
        '''

        pass

    @staticmethod
    @abstractmethod
    def similarity(string1: str, string2: str, ratio: float) -> bool:
        '''
        call `compare` function to compare strings
        compare to ratio and decide whether this count as obfuscation (or not)

        **This function does NOT guarantee similarity based on real meaning of word,
        since this is not semantic segmentation**

        :param string1: string to check obfuscation
        :param string2: another string to check obfuscation
        :param ratio: the ratio for counting obfuscation
        :return: if > ratio return True, count as similar string
        '''