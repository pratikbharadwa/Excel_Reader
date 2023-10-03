import unittest
from openpyxl import Workbook
from Excel_Reader import *

class TestExcelReader(unittest.TestCase):

    def setUp(self):
        self.wb = openpyxl.load_workbook("spain_data.xlsx")
        self.sheet = self.wb["Hoja2"]

    def test_get_start_and_end_rows(self):
    
        start_row, end_row = get_start_and_end_rows(self.sheet)
        self.assertEqual(start_row, 13)
        self.assertEqual(end_row, 14)

    def test_get_start_and_end_columns(self):

        start_col, end_col, weekend_col_start = get_start_and_end_columns(self.sheet)
        self.assertEqual(start_col, 4) 
        self.assertEqual(end_col,103)
        self.assertEqual(weekend_col_start, 54)

    def test_get_time_stamp(self):

        expected_time_stamp=['06.00 a 06.30', '06.30 a 07.00', '07.00 a 07.30', '07.30 a 08.00', '08.00 a 08.30', '08.30 a 09.00', '09.00 a 09.30', '09.30 a 10.00', '10.00 a 10.30', '10.30 a 11.00', '11.00 a 11.30', '11.30 a 12.00', '12.00 a 12.30', '12.30 a 13.00', '13.00 a 13.30', '13.30 a 14.00', '14.00 a 14.30', '14.30 a 15.00', '15.00 a 15.30', '15.30 a 16.00', '16.00 a 16.30', '16.30 a 17.00', '17.00 a 17.30', '17.30 a 18.00', '18.00 a 18.30', '18.30 a 19.00', '19.00 a 19.30', '19.30 a 20.00', '20.00 a 20.30', '20.30 a 21.00', '21.00 a 21.30', '21.30 a 22.00', '22.00 a 22.30', '22.30 a 23.00', '23.00 a 23.30', '23.30 a 24.00', '24.00 a 00.30', '00.30 a 01.00', '01.00 a 01.30', '01.30 a 02.00', '02.00 a 02.30', '02.30 a 03.00', '03.00 a 03.30', '03.30 a 04.00', '04.00 a 04.30', '04.30 a 05.00', '05.00 a 05.30', '05.30 a 06.00', '06.00 a 06.30', '06.30 a 07.00', '07.00 a 07.30', '07.30 a 08.00', '08.00 a 08.30', '08.30 a 09.00', '09.00 a 09.30', '09.30 a 10.00', '10.00 a 10.30', '10.30 a 11.00', '11.00 a 11.30', '11.30 a 12.00', '12.00 a 12.30', '12.30 a 13.00', '13.00 a 13.30', '13.30 a 14.00', '14.00 a 14.30', '14.30 a 15.00', '15.00 a 15.30', '15.30 a 16.00', '16.00 a 16.30', '16.30 a 17.00', '17.00 a 17.30', '17.30 a 18.00', '18.00 a 18.30', '18.30 a 19.00', '19.00 a 19.30', '19.30 a 20.00', '20.00 a 20.30', '20.30 a 21.00', '21.00 a 21.30', '21.30 a 22.00', '22.00 a 22.30', '22.30 a 23.00', '23.00 a 23.30', '23.30 a 24.00', '24.00 a 00.30', '00.30 a 01.00', '01.00 a 01.30', '01.30 a 02.00', '02.00 a 02.30', '02.30 a 03.00', '03.00 a 03.30', '03.30 a 04.00', '04.00 a 04.30', '04.30 a 05.00', '05.00 a 05.30', '05.30 a 06.00']
        start_row, start_col,end_col =13,4,103
        time_stamp=get_time_stamp(self.sheet,start_row,start_col,end_col)
        self.assertEqual(time_stamp,expected_time_stamp)

    def test_get_station_name(self):

        expected_station_name=['TOTAL RADIO DÍA DE AYER']
        start_row, end_row, start_col =13,14,4
        station_name = get_station_name(self.sheet, start_row, end_row, start_col)
        self.assertEqual(station_name,expected_station_name)

    def test_creating_subsets(self):

        expected_result=[{'station_name': 'TOTAL RADIO DÍA DE AYER', 'time': {'06.00': 3960.0, '07.00': 9718.0, '08.00': 13806.0, '09.00': 13137.0, '10.00': 12678.0, '11.00': 11883.0, '12.00': 9979.0, '13.00': 8037.0, '14.00': 6176.0, '15.00': 5268.0, '16.00': 5338.0, '17.00': 6003.0, '18.00': 5966.0, '19.00': 5302.0, '20.00': 4374.0, '21.00': 3339.0, '22.00': 2954.0, '23.00': 3465.0, '24.00': 1719.0, '00.00': 1329.0, '01.00': 1435.0, '02.00': 779.0, '03.00': 579.0, '04.00': 642.0, '05.00': 770.0}, 'Flag': 0}, {'station_name': 'TOTAL RADIO DÍA DE AYER', 'time': {'06.00': 1907.0, '07.00': 4191.0, '08.00': 7149.0, '09.00': 9286.0, '10.00': 11059.0, '11.00': 10843.0, '12.00': 9247.0, '13.00': 6599.0, '14.00': 4300.0, '15.00': 3006.0, '16.00': 3309.0, '17.00': 4170.0, '18.00': 4826.0, '19.00': 4930.0, '20.00': 3762.0, '21.00': 3206.0, '22.00': 2781.0, '23.00': 3206.0, '24.00': 1525.0, '00.00': 1228.0, '01.00': 1414.0, '02.00': 694.0, '03.00': 443.0, '04.00': 445.0, '05.00': 451.0}, 'Flag': 1}]
        start_row, end_row, start_col, weekend_col_start=13,14,4,54
        station_name=['TOTAL RADIO DÍA DE AYER']
        time_stamp=['06.00 a 06.30', '06.30 a 07.00', '07.00 a 07.30', '07.30 a 08.00', '08.00 a 08.30', '08.30 a 09.00', '09.00 a 09.30', '09.30 a 10.00', '10.00 a 10.30', '10.30 a 11.00', '11.00 a 11.30', '11.30 a 12.00', '12.00 a 12.30', '12.30 a 13.00', '13.00 a 13.30', '13.30 a 14.00', '14.00 a 14.30', '14.30 a 15.00', '15.00 a 15.30', '15.30 a 16.00', '16.00 a 16.30', '16.30 a 17.00', '17.00 a 17.30', '17.30 a 18.00', '18.00 a 18.30', '18.30 a 19.00', '19.00 a 19.30', '19.30 a 20.00', '20.00 a 20.30', '20.30 a 21.00', '21.00 a 21.30', '21.30 a 22.00', '22.00 a 22.30', '22.30 a 23.00', '23.00 a 23.30', '23.30 a 24.00', '24.00 a 00.30', '00.30 a 01.00', '01.00 a 01.30', '01.30 a 02.00', '02.00 a 02.30', '02.30 a 03.00', '03.00 a 03.30', '03.30 a 04.00', '04.00 a 04.30', '04.30 a 05.00', '05.00 a 05.30', '05.30 a 06.00', '06.00 a 06.30', '06.30 a 07.00', '07.00 a 07.30', '07.30 a 08.00', '08.00 a 08.30', '08.30 a 09.00', '09.00 a 09.30', '09.30 a 10.00', '10.00 a 10.30', '10.30 a 11.00', '11.00 a 11.30', '11.30 a 12.00', '12.00 a 12.30', '12.30 a 13.00', '13.00 a 13.30', '13.30 a 14.00', '14.00 a 14.30', '14.30 a 15.00', '15.00 a 15.30', '15.30 a 16.00', '16.00 a 16.30', '16.30 a 17.00', '17.00 a 17.30', '17.30 a 18.00', '18.00 a 18.30', '18.30 a 19.00', '19.00 a 19.30', '19.30 a 20.00', '20.00 a 20.30', '20.30 a 21.00', '21.00 a 21.30', '21.30 a 22.00', '22.00 a 22.30', '22.30 a 23.00', '23.00 a 23.30', '23.30 a 24.00', '24.00 a 00.30', '00.30 a 01.00', '01.00 a 01.30', '01.30 a 02.00', '02.00 a 02.30', '02.30 a 03.00', '03.00 a 03.30', '03.30 a 04.00', '04.00 a 04.30', '04.30 a 05.00', '05.00 a 05.30', '05.30 a 06.00']
        subset=creating_subsets(self.sheet,start_row, end_row, start_col, weekend_col_start,time_stamp,station_name)
        self.assertEqual(subset,expected_result)

    def test_insert_data_into_database(self):

        expected=["Insert into radio_audience(station_name, timestamp, audience, flag) values (TOTAL RADIO DÍA DE AYER, {'06.00': 3960.0, '07.00': 9718.0, '08.00': 13806.0, '09.00': 13137.0, '10.00': 12678.0, '11.00': 11883.0, '12.00': 9979.0, '13.00': 8037.0, '14.00': 6176.0, '15.00': 5268.0, '16.00': 5338.0, '17.00': 6003.0, '18.00': 5966.0, '19.00': 5302.0, '20.00': 4374.0, '21.00': 3339.0, '22.00': 2954.0, '23.00': 3465.0, '24.00': 1719.0, '00.00': 1329.0, '01.00': 1435.0, '02.00': 779.0, '03.00': 579.0, '04.00': 642.0, '05.00': 770.0}, 0);"]
        final_audience=[{'station_name': 'TOTAL RADIO DÍA DE AYER', 'time': {'06.00': 3960.0, '07.00': 9718.0, '08.00': 13806.0, '09.00': 13137.0, '10.00': 12678.0, '11.00': 11883.0, '12.00': 9979.0, '13.00': 8037.0, '14.00': 6176.0, '15.00': 5268.0, '16.00': 5338.0, '17.00': 6003.0, '18.00': 5966.0, '19.00': 5302.0, '20.00': 4374.0, '21.00': 3339.0, '22.00': 2954.0, '23.00': 3465.0, '24.00': 1719.0, '00.00': 1329.0, '01.00': 1435.0, '02.00': 779.0, '03.00': 579.0, '04.00': 642.0, '05.00': 770.0}, 'Flag': 0}, {'station_name': 'TOTAL RADIO DÍA DE AYER', 'time': {'06.00': 1907.0, '07.00': 4191.0, '08.00': 7149.0, '09.00': 9286.0, '10.00': 11059.0, '11.00': 10843.0, '12.00': 9247.0, '13.00': 6599.0, '14.00': 4300.0, '15.00': 3006.0, '16.00': 3309.0, '17.00': 4170.0, '18.00': 4826.0, '19.00': 4930.0, '20.00': 3762.0, '21.00': 3206.0, '22.00': 2781.0, '23.00': 3206.0, '24.00': 1525.0, '00.00': 1228.0, '01.00': 1414.0, '02.00': 694.0, '03.00': 443.0, '04.00': 445.0, '05.00': 451.0}, 'Flag': 1}]
        sql_query=insert_data_into_database(final_audience)
        self.assertEqual(expected[0],sql_query)

if __name__ == '__main__':
    unittest.main()