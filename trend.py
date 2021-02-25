import openpyxl
import yahoo_fin
import os


class Trend:
    '''Find probability of price increase after positive and negative trading session'''

    def __init__(self, ticker, after_gain=True, increase=True, append=True, results_path='Results/results.xlsx'):
        self.after_gain = after_gain
        self.increase = True
        self.ticker = ticker
        self.append = append
        self.results_path = results_path

    def get_results(self):
        '''Find probabililties'''
        def count(current_close, previous_close, previous_result):
            '''Adds to counts'''
            if current_close > previous_close:
                gain_days += 1

                if previous_result == 'gain':
                    gain_after_gain += 1
                    return 'gain'
                elif previous_result == 'loss':
                    gain_after_loss += 1
                    return 'gain'

            elif previous_close < current_close:
                loss_days += 1

                return 'loss'

            return

        gain_days = 0
        loss_days = 0
        gain_after_gain = 0
        gain_after_loss = 0

        dummy_list = []
        previous_close = dummy_list[0]
        previous_result = None

        for close in dummy_list[1:]:
            previous_result = count(close, previous_close, previous_result)

        self.gain_after_gain_prob = round(gain_after_gain / gain_days, 2)
        self.gain_after_loss_prob = round(gain_after_loss / loss_days, 2)

    def export_results(self):
        '''Export results into Excel Sheets'''
        def setup_sheets(wb):
            '''Set up freeze panes and colouring'''
            return

        def add_results(wb):
            '''Add results to Excel Sheets'''
            return

        os.makedirs('Results', exist_ok=True)
        if self.append:
            try:
                results_wb = openpyxl.load_workbook(self.results_path)

            except FileNotFoundError:
                print("File Not Found")
        else:
            results_wb = openpyxl.Workbook()

        setup_sheets(results_wb)

        add_results(results_wb)


if __name__ == '__main__':
    tickers = []

    for ticker in tickers:
        stock = Trend(ticker)
        stock.get_results()
        stock.export_results()
