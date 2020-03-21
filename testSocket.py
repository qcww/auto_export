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
    post_data = "2019年 06月份 进项票勾选认证数据导出"
    ws.send('{"type":"action","device_client_name":"fd657b34-5ec4-11ea-a813-005056c00008","room_id":"invoice_user_group","data":"%s"}' % post_data)

websocket.enableTrace(True)
ws = websocket.WebSocketApp("ws://im.itking.cc:12366",
                            on_message=on_message,
                            on_error=on_error,
                            on_close=on_close)
ws.on_open = on_open

ws.run_forever()