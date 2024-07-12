import xlwings as xw
from alice_credentials import*
import pandas as pd



alice = login()

# Excel:
wb = xw.Book('Excel_Based_Algo_Trading_System.xlsx')
sht = wb.sheets['Sheet1']
sht.range('C1:J200').value = None
sht.range('M1:T1').value = 'Product Type(Intraday/Delivery)',"Direction(Buy/Sell)","Qty","Trigger Price","Limit Price", "Entry","Order(Modify/Cancel)","Order_Status"

# WebSocket Connection 
LTP = 0
socket_opened = False
subscribe_flag = False
subscribe_list = []
unsubscribe_list = []
data = {}

def socket_open():
    print("Connected")
    global socket_opened
    socket_opened = True
    if subscribe_flag:
        alice.subscribe(subscribe_list)

def socket_close():
    global socket_opened, LTP
    socket_opened = False
    LTP = 0
    print("Closed")

def socket_error(message):
    global LTP
    LTP = 0
    print("Error :", message)

def feed_data(message):
    global LTP, subscribe_flag, data
    feed_message = json.loads(message)
    if feed_message["t"] == "ck":
        print("Connection Acknowledgement status :%s (Websocket Connected)" % feed_message["s"])
        subscribe_flag = True
        print("subscribe_flag :", subscribe_flag)
        print("-------------------------------------------------------------------------------")
        pass
    elif feed_message["t"] == "tk":
        token = feed_message["tk"]
        if "ts" in feed_message:
            symbol = feed_message["ts"]
        else:   
            symbol = token  # For indices
        data[symbol] = {
            "Open": feed_message.get("o", 0),
            "High": feed_message.get("h", 0),
            "Low": feed_message.get("l", 0),
            "LTP": feed_message.get("lp", 0),
            "OI": feed_message.get("toi", 0),
            "VWAP": feed_message.get("ap", 0),
            "PrevDayClose": feed_message.get("c", 0),
                   }
        # print(f"Token Acknowledgement status for {symbol}: {feed_message}")
        print("-------------------------------------------------------------------------------")
        pass
    else:
        # print("Feed :", feed_message)
        LTP = feed_message["lp"] if "lp" in feed_message else LTP

alice.start_websocket(socket_open_callback=socket_open, socket_close_callback=socket_close,
                      socket_error_callback=socket_error, subscription_callback=feed_data, run_in_background=True, market_depth=False)


def algo():
    for row_no in range(2,200) :
        
        productType = sht.range('M'+str(row_no)).value
        direction = sht.range('N'+str(row_no)).value
        qty = sht.range('O'+str(row_no)).value
        Trigger_Price = sht.range('P'+str(row_no)).value
        Limit_Price = sht.range('Q'+str(row_no)).value
        entry = sht.range('R'+str(row_no)).value
        Order = sht.range('S'+str(row_no)).value
        Order_Status = sht.range('T'+str(row_no)).value
        
        if entry == "True" and Order_Status is not None:
            
            exchange = sht.range('A'+str(row_no)).value
            symbol = sht.range('B'+str(row_no)).value
            print ("%%%%%%%%%%%%%%%%%%%%%%%%%%%%1%%%%%%%%%%%%%%%%%%%%%%%%%%%%%")
            print(
            alice.place_order(transaction_type = TransactionType.Buy if direction == "Buy" else TransactionType.Sell,
                    instrument = alice.get_instrument_by_symbol(exchange,symbol),
                    quantity = int(qty),
                    order_type = OrderType.StopLossMarket,
                    product_type = ProductType.Delivery if productType == "Delivery" else ProductType.Intraday,
                    price = 0.0,
                    trigger_price = None,
                    stop_loss = None,
                    square_off = None,
                    trailing_sl = None,
                    is_amo = False,
                    order_tag='order1')
            )
            print("Order Placed")
            # Order_Status = "Placed Order"
            # sht.range('T'+str(row_no)).value = Order_Status



instrument = []
for row in sht.range("A2:B200").value:
    exchange,symbol = row
    if exchange and symbol:
        instrument.append((exchange,symbol))

subscribe_list =[]
for exchange,symbol in instrument:
    subscribe_list.append(alice.get_instrument_by_symbol(exchange,symbol))
    new_subscribe_list = subscribe_list


while True:
    alice.subscribe(new_subscribe_list)
    #DataFrame:
    df = pd.DataFrame.from_dict(data,orient='index')
    sht.range("C1").value = df 
    print(df)   

    algo()

