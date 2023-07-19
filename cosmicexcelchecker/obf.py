# script related to obfuscation (fuzzy word)

from cosmicexcelchecker._baseclass import AbstractObf

class CheckObf(AbstractObf):
    '''
    Concrete implementation of the AbstractObf class
    Provide static method to compare two strings
    '''

    @staticmethod
    def compare(string1: str, string2: str) -> int:
        '''
        Compare the edit distance of two given strings.
        Use the Levenshtein distance algorithem, which allows edit, add, delete
        Calculate the step numbers

        :param string1: a string input
        :param string2: another string input
        :return: step numbers as int
        '''

        # use var to get string1 + 1and string2 length + 1 since a lot of usage later
        len_str1_add1 = len(string1) + 1
        len_str2_add1 = len(string2) + 1

        # create mem space for dp, 1D array for performance but think as 2d array
        len_dp_mem = len_str1_add1 * len_str2_add1
        dp_mem = [0 for _ in range(len_dp_mem)]

        # generate for (a_i, 0) and (0, b_j) cells
        for i in range(len_str1_add1):
            dp_mem[i] = i

        for j in range(0, len_dp_mem, len_str1_add1):
            dp_mem[j] = j // len_str1_add1

        # iterate through each cell, check the min based on 3 ways (see levenshtein distance formula)
        for i in range(1, len_dp_mem):  # skip first zero to improve perf
            if dp_mem[i] != 0:  # first line or first col in 2D
                continue

            dp_mem[i] = min(dp_mem[i - 1] + 1, dp_mem[i - len_str1_add1] + 1, dp_mem[i - len_str1_add1 - 1] + \
                int(string1[i % len_str1_add1 - 1] != string2[i // len_str1_add1 - 1]))

        return dp_mem[-1]  # return the bottom right if in 2D

    @staticmethod
    def similarity(string1: str, string2: str, ratio: float) -> bool:
        '''
        use compare() to generate distance, "(longer - compare()) / longer" to compare ratio
        If greater than ratio, means similar. Otherwise, it's not
        Use 2-digit rounding, meaning num >= 0.785 -> 0.79

        :param string1: First string input
        :param string2: Second string input
        :param ratio: Given ratio to compare, default to 0.79
        :return: True if similar, False if not similar by comparing the ratio
        '''

        ed : int = CheckObf.compare(string1=string1, string2=string2)

        return round((max(len(string1), len(string2)) - ed) / max(len(string1), len(string2)), 2) >= ratio



