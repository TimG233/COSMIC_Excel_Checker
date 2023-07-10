# Custom error class

class CosmicExcelCheckerException(Exception):
    '''
    Base Exception Class for the module
    '''

    pass

class IncorrectFileTypeException(CosmicExcelCheckerException):
    '''
    Raised when received incorrect file type. E.g: try to use load_excel to load files that are not excel.
    '''

    pass

class UnknownREQNumException(CosmicExcelCheckerException):
    '''
    Raised if a single result summary file does not have specific req number
    '''

    pass

class RepeatedREQNumException(CosmicExcelCheckerException):
    '''
    Raised if a single result summary file has repeated requirement number
    '''

    pass

class SheetNotFoundException(CosmicExcelCheckerException):
    '''
    Raised if a sheet is not found by given name
    '''

    pass