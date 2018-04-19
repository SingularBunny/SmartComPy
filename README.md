# SmartComPy
Module provide access to SmartCOM API of versions 3 and 4 from ITI Capital (IT Invest).
The script can be used in two modes:
-  server mode
-  module mode

## Server mode
Allow to split module of your system in two separate parts. First part is a Windows servev with SmartCOM API and wrapper script. Second part is you code in under Linux or Windows. In server mode your could run script as server(under Windows only):
```shell
python smartcom3.py
```
Then from your code (Linux or Windows Machine) connect to
remote manager (server) like this:

```python
# define manager client
class SmartCOM3Manager(BaseManager): 
	pass

# register necessarry methods
SmartCOM3Manager.register('get_smartcom3_server')
SmartCOM3Manager.register('get_smartcom3_event_queue')

# connect
m = SmartCOM3Manager(address=('foo.bar.org', 50000), authkey='abracadabra')
m.connect()

# get SmartCom server and event queue instances
server = m.get_smartcom3_server()
current_process().authkey = 'abracadabra' #should be the same as manager's authkey
event_queue = server.get_event_queue()

# and use SmartCOM3 API.
server.connect('server', 'port', 'login', 'password')
```

## Module mode (Windows only)
Just import SmartCOM3Manager and use.

```python
from smartcom3 import SmartCOM3Manager

manager = SmartCOM3Manager()
manager.start()

smartcom3_server = manager.get_smartcom3_server()
event_queue = server.get_event_queue()
```

Method `get_smartcom3_server()` returns SmartCOM server from API.
Methods GetBarsSer() and GetTradesSer() are needed to support standart python datetime type.
Events come trough event queue as tuples.
