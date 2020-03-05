#coding=utf-8

import websocket
import json

def on_message(ws, message):
    print(message)
    
    


def on_error(ws, error):
    print(error)


def on_close(ws):
    print("### closed ###")

def on_open(ws):
    ws.send('{"type":"login","room_id":"invoice_user_group","client_name":"%s"}' % 'tax_ah4')
    post_data = "2019年 11月份 航天销项票数据导出"
    ws.send('{"type":"action","device_client_name":"07fcbd2c-5c52-11ea-9d2b-005056c00008","room_id":"invoice_user_group","data":"%s"}' % post_data)

websocket.enableTrace(True)
ws = websocket.WebSocketApp("ws://im.itking.cc:12366",
                            on_message=on_message,
                            on_error=on_error,
                            on_close=on_close)
ws.on_open = on_open

ws.run_forever()