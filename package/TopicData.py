import openpyxl
import copy
import re

TITLE, ID, START_TIME, END_TIME = range(4)

class TopicData:
    """
    TopicData object contains the information in a particular row

    Arguments:
    None

    Container: dictionary
    key --> column number
    value --> list(   TITLE, (ID, (START_TIME, END_TIME))   ) 

    Functions:
    add() : add the video information of one cell (in list form) into the container. 
    empty() : empty the container
    split() : split the raw time input into start time and end time 
    """

    def __init__(self):
        self.m_dict = {}

    def __getitem__(self, key):
        return self.m_dict[key]

    def __len__(self):
        return len(self.m_dict)

    def __str__(self):
        return str(self.m_dict)

    def __repr__(self):
        return self.__str__()

    # implement deepcopy to avoid aliasing when the TopicData instance is
    # appended into VideoData instance
    def __deepcopy__(self, memo):
        return copy.deepcopy(self.m_dict)

    def add(self, index, list_var):
        """
        Append the information in a cell into the m_dict.

        Arguments: 
        self --- TopicData object
        index (key) --- indexing number, starts from 1
        list_var (value) --- a list of 3 things in each column: title, iD, raw time

        Return: 
        None
        """
        self.m_dict[index] = list_var

    def empty(self):
        """empty the m_dict """
        self.m_dict = {}

    def isEmpty(self):
        """return True if the m_dict is empty"""
        return not len(self.m_dict)

    def splitTime(self):
        """
        split the raw time into start time and end time

        return: None

        """
        for i in range(1, len(self.m_dict) + 1):

            split_time_list = []
            split_time_list.append(self.m_dict[i][TITLE])
            split_time_list.append(self.m_dict[i][ID])

            # start from 2 to ignore the title and the first ID
            for raw_time_input in range(2, len(self.m_dict[i])):

                temp_str = self.m_dict[i][raw_time_input]

                # if the raw time is 'whole' --> split into '' and ''
                if temp_str == 'whole':
                    split_time_list.append('')
                    split_time_list.append('')

                # split raw time into start time and end time
                else:
                    if self.isTime(temp_str) is False:
                        split_time_list.append(temp_str)  # put ID back

                    else:
                        temp_list = temp_str.split('-')
                        for j in range(0, len(temp_list)):
                            if temp_list[j] == 'start' or temp_list[j] == 'end':
                                temp_list[j] = ''
                        split_time_list.append(temp_list[0])
                        split_time_list.append(temp_list[1])
            self.m_dict[i] = split_time_list

    def isTime(self, string):
        """
        check if the string is ID or raw time

        Arguments:
        self --- TopicData
        string --- the information in the cell which has been stored in m_dict value

        Return:
        bool --- True if raw time, False if ID
        """
        temp_list = string.split('-')

        # ID may contain hyphen
        if len(temp_list) == 2:
            # raw time consists of actual time, start, or end
            if temp_list[0] == 'start' or temp_list[1] == 'end':
                return True
            elif re.search(":", string):
                return True
            else:
                return False
        else:
            # for case where ID doesn't contain hyphen
            return False