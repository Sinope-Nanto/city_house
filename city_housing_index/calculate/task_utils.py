from utils.calculate_utils import linearRegression, get_ratio
import numpy as np


class DataPreProcessor:

    @staticmethod
    def transfer(data):
        return np.array(data).T


class ModelCalculateUtils:
    def __init__(self, input):
        self.__data = DataPreProcessor.transfer(input)

    def get_result(self):
        return linearRegression(self.__data)

    def review_data(self):
        return self.__data


class PriceSequenceUtils:
    def __init__(self, input_current, input_last_month, input_last_year):
        self.__data_current = DataPreProcessor.transfer(input_current)
        self.__data_last_month = DataPreProcessor.transfer(input_last_month)
        self.__data_last_year = DataPreProcessor.transfer(input_last_year)

    def get_result(self):
        return get_ratio(self.__data_current, self.__data_last_month, self.__data_last_year)
