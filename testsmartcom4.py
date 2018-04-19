import Queue
import sys
import time
import unittest
from multiprocessing import current_process
from subprocess import Popen
from unittest import TestCase

from yaml import load

from trader.core.smartcom4 import SmartCOM4Manager

try:
    from yaml import CLoader as Loader, CDumper as Dumper
except ImportError:
    from yaml import Loader, Dumper

class IterableQueue():
    def __init__(self,source_queue):
            self.source_queue = source_queue
    def __iter__(self):
        while True:
            try:
               yield self.source_queue.get_nowait()
            except Queue.Empty:
               return

class TestSmartcom4(TestCase):
    def setUp(self):
        super(TestSmartcom4, self).setUp()
        self.smartcom4 = Popen([sys.executable, 'trader/core/smartcom4.py'])
        with open('configuration/config-client.yaml', 'r') as confFile:
            config = load(confFile, Loader=Loader)
            srv_conf = config.get('pythonServer')
            self.clnt_conf = config.get('application')
            self.manager = SmartCOM4Manager(address=(srv_conf.get('address'), srv_conf.get('port')),
                                            authkey=srv_conf.get('authkey'))
            self.manager.connect()
            self.smartcom4_server = self.manager.get_smartcom4_server()
            current_process().authkey = srv_conf.get('authkey')

    def tearDown(self):
        super(TestSmartcom4, self).tearDown()
        self.smartcom4.terminate()

    def test_connection(self):
        self.smartcom4_server.connect(self.clnt_conf.get('server'), self.clnt_conf.get('port'),
                                      self.clnt_conf.get('login'), self.clnt_conf.get('password'))
        event_queue = self.manager.get_event_queue()
        it_queue = IterableQueue(event_queue)

        time.sleep(5)
        self.smartcom4_server.PlaceOrder()
        self.smartcom4_server.disconnect()
        time.sleep(2)

        self.assertEquals(2, event_queue.qsize())
        self.assertEquals('Connected', event_queue.get_nowait()[0])
        self.assertEquals('Disconnected', event_queue.get_nowait()[0])

if __name__ == '__main__':
    unittest.main()
