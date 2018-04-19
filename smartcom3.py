"""
Module provide access to SmartCOM3 API.
"""
import sys

import logging

sys.coinit_flags = 0

import time
import Queue
from datetime import datetime as dt
from multiprocessing import current_process, freeze_support
from multiprocessing.managers import BaseManager

try:
    from servicemanager import CoInitializeEx, COINIT_MULTITHREADED, CoUninitialize
    import pywintypes
    from win32com import client
except ImportError:
    pass

try:
    from yaml import CLoader as Loader, CDumper as Dumper, load
except ImportError:
    from yaml import Loader, Dumper, load

logger = logging.getLogger()
handler = logging.StreamHandler()
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

client_events = ['OrderFailed',
                 'SetMyClosePos',
                 'SetMyTrade',
                 'UpdateBidAsk',
                 'SetMyOrder',
                 'AddTrade',
                 'SetSubscribtionCheckReult',
                 'OrderMoveSucceeded',
                 'SetPortfolio',
                 'Connected',
                 'UpdatePosition',
                 'Disconnected',
                 'AddTick',
                 'OrderCancelFailed',
                 'OrderMoveFailed',
                 'OrderSucceeded',
                 'UpdateOrder',
                 'AddTickHistory',
                 'OrderCancelSucceeded',
                 'AddBar',
                 'UpdateQuote',
                 'AddPortfolio',
                 'AddSymbol']


class Constants:
    StOrder_State_Cancel = 4  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_ContragentCancel = 7  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_ContragentReject = -1  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_Expired = 3  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_Filled = 5  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_Open = 2  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_Partial = 6  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_Pending = 1  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_Submited = 0  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_SystemCancel = 9  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_State_SystemReject = 8  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0001
    StOrder_Action_Buy = 1  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0002
    StOrder_Action_Cover = 4  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0002
    StOrder_Action_Sell = 2  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0002
    StOrder_Action_Short = 3  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0002
    StOrder_Type_Limit = 2  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0003
    StOrder_Type_Market = 1  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0003
    StOrder_Type_Stop = 3  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0003
    StOrder_Type_StopLimit = 4  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0003
    StOrder_Validity_Day = 1  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0004
    StOrder_Validity_Gtc = 2  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0004
    StBarInterval_10Min = 3  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_15Min = 4  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_1Min = 1  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_2Hour = 7  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_30Min = 5  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_4Hour = 8  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_5Min = 2  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_60Min = 6  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_Day = 9  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_Month = 11  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_Quarter = 12  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_Tick = 0  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_Week = 10  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StBarInterval_Year = 13  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0005
    StPortfolioStatus_AutoRestricted = 5  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0006
    StPortfolioStatus_Blocked = 3  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0006
    StPortfolioStatus_Broker = 0  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0006
    StPortfolioStatus_OrderNotSigned = 6  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0006
    StPortfolioStatus_ReadOnly = 2  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0006
    StPortfolioStatus_Restricted = 4  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0006
    StPortfolioStatus_TrustedManagement = 1  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0006
    StErrorCode_BadParameters = -1610612732  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0007
    StErrorCode_ExchangeNotAccessible = -1610612730  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0007
    StErrorCode_InternalError = -1610612731  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0007
    StErrorCode_NotConnected = -1610612733  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0007
    StErrorCode_PortfolioNotFound = -1610612734  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0007
    StErrorCode_SecurityNotFound = -1610612735  # from enum __MIDL___MIDL_itf_SmartCOM3_0000_0000_0007

queue = Queue.Queue()
class StClientEvents:
    def __init__(self):
        self.event_queue = queue

    def OnOrderFailed(self, cookie, orderid, reason):
        self.event_queue.put_nowait(('OrderFailed', cookie, orderid, reason))

    def OnSetMyClosePos(self, row, nrows, portfolio, symbol, amount, price_buy, price_sell, postime, order_open,
                        order_close):
        self.event_queue.put_nowait(('SetMyClosePos', row, nrows, portfolio, symbol, amount, price_buy, price_sell,
                                     pytime_2_datetime(postime), order_open, order_close))

    def OnSetMyTrade(self, row, nrows, portfolio, symbol, datetime, price, volume, tradeno, buysell, orderno):
        self.event_queue.put_nowait(
            ('SetMyTrade', row, nrows, portfolio, symbol, pytime_2_datetime(datetime), price, volume, tradeno, buysell,
             orderno))

    def OnUpdateBidAsk(self, symbol, row, nrows, bid, bidsize, ask, asksize):
        self.event_queue.put_nowait(('UpdateBidAsk', symbol, row, nrows, bid, bidsize, ask, asksize))

    def OnSetMyOrder(self, row, nrows, portfolio, symbol, state, action, type, validity, price, amount, stop, filled,
                     datetime, id, no, cookie):
        self.event_queue.put_nowait(('SetMyOrder', row, nrows, portfolio, symbol, state, action, type, validity, price,
                                     amount, stop, filled, pytime_2_datetime(datetime), id, no, cookie))

    def OnAddTrade(self, portfolio, symbol, orderid, price, amount, datetime, tradeno):
        self.event_queue.put_nowait(
            ('AddTrade', portfolio, symbol, orderid, price, amount, pytime_2_datetime(datetime), tradeno))

    def OnSetSubscribtionCheckReult(self, result):
        self.event_queue.put_nowait(('SetSubscribtionCheckReult', result))

    def OnOrderMoveSucceeded(self, orderid):
        self.event_queue.put_nowait(('OrderMoveSucceeded', orderid))

    def OnSetPortfolio(self, portfolio, cash, leverage, comission, saldo):
        self.event_queue.put_nowait(('SetPortfolio', portfolio, cash, leverage, comission, saldo))

    def OnConnected(self):
        self.event_queue.put_nowait(('Connected',))

    def OnUpdatePosition(self, portfolio, symbol, avprice, amount, planned):
        self.event_queue.put_nowait(('UpdatePosition', portfolio, symbol, avprice, amount, planned))

    def OnDisconnected(self, reason):
        self.event_queue.put_nowait(('Disconnected', reason))

    def OnAddTick(self, symbol, datetime, price, volume, tradeno, action):
        self.event_queue.put_nowait(('AddTick', symbol, pytime_2_datetime(datetime), price, volume, tradeno, action))

    def OnOrderCancelFailed(self, orderid):
        self.event_queue.put_nowait(('OrderCancelFailed', orderid))

    def OnOrderMoveFailed(self, orderid):
        self.event_queue.put_nowait(('OrderMoveFailed', orderid))

    def OnOrderSucceeded(self, cookie, orderid):
        self.event_queue.put_nowait(('OrderSucceeded', cookie, orderid))

    def OnUpdateOrder(self, portfolio, symbol, state, action, type, validity, price, amount, stop, filled, datetime,
                      orderid, orderno, status_mask, cookie):
        self.event_queue.put_nowait(('UpdateOrder', portfolio, symbol, state, action, type, validity, price, amount,
                                     stop, filled, pytime_2_datetime(datetime), orderid, orderno, status_mask, cookie))

    def OnAddTickHistory(self, row, nrows, symbol, datetime, price, volume, tradeno, action):
        self.event_queue.put_nowait(
            ('AddTickHistory', row, nrows, symbol, pytime_2_datetime(datetime), price, volume, tradeno, action))

    def OnOrderCancelSucceeded(self, orderid):
        self.event_queue.put_nowait(('OrderCancelSucceeded', orderid))

    def OnAddBar(self, row, nrows, symbol, interval, datetime, open, high, low, close, volume, open_int):
        self.event_queue.put_nowait(
            ('AddBar', row, nrows, symbol, interval, pytime_2_datetime(datetime), open, high, low, close, volume,
             open_int))

    def OnUpdateQuote(self, symbol, datetime, open, high, low, close, last, volume, size, bid, ask, bidsize, asksize,
                      open_int, go_buy, go_sell, go_base, go_base_backed, high_limit, low_limit, trading_status, volat,
                      theor_price):
        self.event_queue.put_nowait(
            ('UpdateQuote', symbol, pytime_2_datetime(datetime), open, high, low, close, last, volume, size, bid,
             ask, bidsize, asksize, open_int, go_buy, go_sell, go_base, go_base_backed, high_limit, low_limit,
             trading_status, volat, theor_price))

    def OnAddPortfolio(self, row, nrows, portfolioName, portfolioExch, portfolioStatus):
        self.event_queue.put_nowait(('AddPortfolio', row, nrows, portfolioName, portfolioExch, portfolioStatus))

    def OnAddSymbol(self, row, nrows, symbol, short_name, long_name, type, decimals, lot_size, punkt, step, sec_ext_id,
                    sec_exch_name, expiry_date, days_before_expiry, strike):
        self.event_queue.put_nowait(('AddSymbol', row, nrows, symbol, short_name, long_name, type, decimals, lot_size,
                                     punkt, step, sec_ext_id, sec_exch_name, pytime_2_datetime(expiry_date),
                                     days_before_expiry, strike))

    # GetBars method that accept serializable datetime object instead of PyTime object.
    def GetBarsSer(self, symbol, interval, since, count):
        self.GetBars(symbol, interval, datetime_2_pytime(since), count)

    # GetTrades method that accept serializable datetime instead of PyTime object.
    def GetTradesSer(self, symbol, from_, count):
        self.GetTrades(symbol, datetime_2_pytime(from_), count)


def pytime_2_datetime(pytime):
    return dt(year=pytime.year,
              month=pytime.month,
              day=pytime.day,
              hour=pytime.hour,
              minute=pytime.minute,
              second=pytime.second,
              microsecond=pytime.msec)


def datetime_2_pytime(datetime):
    return pywintypes.Time(time.mktime(datetime.timetuple()))


class SmartCOM3Manager(BaseManager):
    pass


def get_smartcom3_server():
    CoInitializeEx(COINIT_MULTITHREADED)
    clnt = client.DispatchWithEvents('SmartCOM3.StServer.1', StClientEvents)
    CoUninitialize()
    return clnt

def get_event_queue():
    return queue

SmartCOM3Manager.register('get_event_queue', callable=get_event_queue)
SmartCOM3Manager.register('get_smartcom3_server',
                          callable=get_smartcom3_server,
                          exposed=('CancelBidAsks', 'CancelOrder', 'CancelPortfolio', 'CancelQuotes', 'CancelTicks',
                                   'ConfigureClient', 'ConfigureServer', 'GetBars', 'GetMoneyAccount', 'GetMyClosePos',
                                   'GetMyOrders', 'GetMyTrades', 'GetPrortfolioList', 'GetSymbols', 'GetTrades',
                                   'IsConnected', 'ListenBidAsks', 'ListenPortfolio', 'ListenQuotes', 'ListenTicks',
                                   'MoveOrder', 'PlaceOrder', 'connect', 'disconnect', 'GetBarsSer',
                                   'GetTradesSer'))

if __name__ == '__main__':
    freeze_support()

    from argparse import ArgumentParser, FileType

    parser = ArgumentParser(description='python wraper for SmartCOM3')

    parser.add_argument('-c', '--config',
                        type=FileType('r', 0),
                        dest='config',
                        default='configuration/config-server.yaml',
                        metavar='configuration/config-server.yaml',
                        help='path to configuration file')

    args = parser.parse_args()

    with args.config as confFile:
        config = load(confFile, Loader=Loader).get('application')
        current_process().authkey = config.get('authkey')
        logger.setLevel(config['logger']['level'])
        logger.info('SmartCOM3 Server running...')
        m = SmartCOM3Manager(address=('', config.get('port')),
                             authkey=config.get('authkey')).get_server().serve_forever()
