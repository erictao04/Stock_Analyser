import openpyxl
import yfinance as yf
import os


class Trend:
    '''Find probability of price increase after positive and negative trading session'''

    def __init__(self, ticker, after_gain=True, increase=True, append=True, results_path='Results/results.xlsx'):
        self.after_gain = after_gain
        self.increase = True
        self.ticker = ticker.upper()
        self.append = append
        self.results_path = results_path

    def get_data(self):
        self.closes = yf.Ticker(self.ticker).history(period="max")['Close']

    def get_results(self):
        '''Find probabililties'''
        def count(current_close, previous_close, previous_result):
            '''Adds to counts'''
            if current_close > previous_close:
                self.gain_days += 1

                if previous_result == 'gain':
                    self.gain_after_gain += 1
                    return 'gain'
                elif previous_result == 'loss':
                    self.gain_after_loss += 1
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
    tickers = ['AC', 'ADP', 'AEE', 'AGI', 'AI']

    for ticker in tickers:
        stock = Trend(ticker)
        stock.get_data()
        stock.get_results()
        print(
            f'''After gain: {stock.gain_after_gain_prob}%, After loss: {stock.gain_after_loss_prob}%''')
#        stock.export_results()
