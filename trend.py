import openpyxl
import yfinance as yf
import os
from pandas_datareader import data as pdr
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter


class Trend:
    '''Find probability of price increase after positive and negative trading session'''

    def __init__(self, ticker, after_gain=True, increase=True, auto_append=False, append=True, results_path='Results/results.xlsx'):
        self.after_gain = after_gain
        self.increase = True
        self.ticker = ticker.upper()
        self.append = append
        self.results_path = results_path
        self.auto_append = auto_append

    def get_data(self):
        yf.pdr_override()
        self.closes = pdr.get_data_yahoo(self.ticker)['Close']

    def get_results(self):
        '''Find probabililties'''
        def count(current_close, previous_close, previous_result):
            '''Adds to counts'''
            if current_close > previous_close:
                self.gain_days += 1

                if previous_result == 'gain':
                    self.gain_after_gain += 1
                elif previous_result == 'loss':
                    self.gain_after_loss += 1
                return 'gain'

            elif current_close < previous_close:
                self.loss_days += 1

                return 'loss'

            return

        def count2(current_close, previous_close, previous_result):
            if previous_result == 'gain':
                if current_close > previous_close:
                    self.gain_after_gain += 1
                    self.gain_days += 1
                    return 'gain'
                elif current_close < previous_close:
                    self.loss_days += 1
                    return 'loss'

            elif previous_result == 'loss':
                if current_close > previous_close:
                    self.gain_after_loss += 1
                    self.gain_days += 1
                    return 'gain'
                elif current_close < previous_close:
                    self.loss_days += 1
                    return 'loss'

            else:
                if current_close > previous_close:
                    self.gain_days += 1
                    return 'gain'
                elif current_close < previous_close:
                    self.loss_days += 1
                    return 'loss'

            return

        self.gain_days, self.loss_days, self.gain_after_gain, self.gain_after_loss = 0, 0, 0, 0

        previous_close = self.closes[0]
        previous_result = None

        for (date, close) in self.closes[1:].iteritems():
            previous_result = count(close, previous_close, previous_result)
            previous_close = close

        self.gain_after_gain_prob = round(
            self.gain_after_gain / self.gain_days, 2)
        self.gain_after_loss_prob = round(
            self.gain_after_loss / self.loss_days, 2)

    def export_results(self):
        '''Export results into Excel Sheets'''
        def setup_sheets(wb):
            '''Set up freeze panes and colouring'''
            results_sheet = wb['Sheet']
            results_sheet.freeze_panes = 'B2'
            results_sheet['A1'] = 'Ticker'
            results_sheet['B1'] = 'After Gains (%)'
            results_sheet['C1'] = 'After Loss (%)'

            for i in range(1, 4):
                letter = get_column_letter(i)
                results_sheet[f'{letter}1'].alignment = centered
                results_sheet.column_dimensions['B'].width = 15
                results_sheet.column_dimensions['C'].width = 15

            return results_sheet

        def add_results(sheet):
            '''Add results to Excel Sheets'''
            new_row = sheet.max_row + 1
            sheet[f'A{new_row}'] = self.ticker
            sheet[f'B{new_row}'] = self.gain_after_gain_prob
            sheet[f'C{new_row}'] = self.gain_after_loss_prob

            sheet[f'A{new_row}'].alignment = centered
            sheet[f'B{new_row}'].alignment = centered
            sheet[f'C{new_row}'].alignment = centered

            if new_row % 2 != 0:
                sheet[f'A{new_row}'].fill = grey_fill
                sheet[f'B{new_row}'].fill = grey_fill
                sheet[f'C{new_row}'].fill = grey_fill

            return sheet

        centered = Alignment(horizontal='center')
        grey_fill = PatternFill("solid", fgColor='00C0C0C0')

        os.makedirs('Results', exist_ok=True)

        if self.auto_append:
            try:
                results_wb = openpyxl.load_workbook(self.results_path)
                results_sheet = results_wb['Sheet']
            except FileNotFoundError:
                results_wb = openpyxl.Workbook()
                results_sheet = setup_sheets(results_wb)
        elif self.append:
            try:
                results_wb = openpyxl.load_workbook(self.results_path)
                results_sheet = results_wb['Sheet']

            except FileNotFoundError:
                print("File Not Found")
                return
        else:
            results_wb = openpyxl.Workbook()
            results_sheet = setup_sheets(results_wb)

        results_sheet = add_results(results_sheet)

        results_wb.save(self.results_path)


if __name__ == '__main__':
    import csv
    with open('Tickers2.csv') as csv_file:
        csv_reader = csv.reader(csv_file)

        for ticker in csv_reader:
            if "." in ticker[0]:
                continue
            stock = Trend(ticker[0], auto_append=True)
            stock.get_data()
            stock.get_results()
            stock.export_results()
