from lib.DataProcessor import DataProcessor
from lib.ExcelManipulator import ExcelManipulator
from lib.GUIApp import GUIApp

def main():
    # data_processor = DataProcessor()
    # excel_manipulator = ExcelManipulator('/Users/ayoub/Documents/goaluin/data-duplication-solution/tests/Example 1.xlsx')

    # data = excel_manipulator.extract_data()

    # grouped_data = data_processor.group_duplicates(data, [0, 2], [0, 2])
    # merged_data = data_processor.merge_data(grouped_data, [0, 2])

    # for row in merged_data:
    #     print(row)

    app = GUIApp()
    app.start()

main()

