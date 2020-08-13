import pandas as pd
from pandas import DataFrame
import weakref


df = pd.read_excel('stock.xlsx')
inventory_code = df['code'].tolist()
inventory_code2 = df['code'].tolist()
inventory_name = df['name'].tolist()
inventory_amount = df['amount'].tolist()
inventory_unit = df["unit"].tolist()
inventory_location = df["location"].tolist()



def reload():
    global inventory_name, inventory_amount_new, inventory_unit_new, inventory_code, inventory_location_new,\
        inventory_name, inventory_amount, inventory_unit, inventory_code, inventory_location
    df=pd.read_excel("stock.xlsx")
    inventory_code = df['code'].tolist()
    inventory_name = df['name'].tolist()
    inventory_amount = df['amount'].tolist()
    inventory_unit = df["unit"].tolist()
    inventory_location = df["location"].tolist()





df3 = pd.read_excel('OrderCurrent.xlsx')

OrderCurrent_OrderNo = df3['OrderCurrent_OrderNo'].tolist()
OrderCurrent_ProductCode = df3["OrderCurrent_ProductCode"].tolist()
OrderCurrent_ProductCode2 = df3["OrderCurrent_ProductCode"].tolist()
OrderCurrent_ProductName = df3["OrderCurrent_ProductName"].tolist()
OrderCurrent_Amount = df3["OrderCurrent_Amount"].tolist()
OrderCurrent_Date = df3["OrderCurrent_Date"].tolist()

inventory_name_new = []
inventory_amount_new = []
inventory_unit_new = []
inventory_location_new = []



class InventoryClass(object):

    def __init__(self, i_code, i_amount, i_unit, i_location):
        self.i_code = i_code
        self.i_amount = i_amount
        self.i_unit = i_unit
        self.i_location = i_location

    def info(self):
        print("Product code:", self.i_code, "Amount:", self.i_amount, "Unit:", self.i_unit, "Place:", self.i_location)

    def add(self, add_amount):
        self.i_amount += add_amount

    def subtract(self, subtract_amount):
        self.i_amount -= subtract_amount

    def change_location(self, new_location):
        self.i_location = new_location

    def change_unit(self, new_unit):
        self.i_unit = str(new_unit)

    def amount(self):
        return self.i_amount

    def update(self):
        reload()
        for i in range(0, len(inventory_name)):
            # if "a"=="a":
            if inventory_code[i] == self.i_code:
                inventory_unit_new.insert(i, self.i_unit)
                inventory_amount_new.insert(i, self.i_amount)
                inventory_location_new.insert(i, self.i_location)
            else:
                inventory_unit_new.insert(i, inventory_unit[i])
                inventory_amount_new.insert(i, inventory_amount[i])
                inventory_location_new.insert(i, inventory_location[i])

        p = zip(inventory_name, inventory_amount_new, inventory_unit_new, inventory_code, inventory_location_new)
        df2 = DataFrame(p)
        df2.columns = ["name", "amount", "unit", "code", "location"]
        writer = pd.ExcelWriter('stock.xlsx', engine='xlsxwriter')
        df2.to_excel(writer, sheet_name='Sheet1')
        writer.save()
        writer.close()




for i in range(0, len(inventory_name)):
    globals()[inventory_name[i]] = InventoryClass(inventory_code[i], inventory_amount[i],
                                                  inventory_unit[i], inventory_location[i])


class InventoryListClass(object):
    _instances = set()

    def __init__(self, inventory_code, inventory_name, inventory_amount, inventory_unit, inventory_location):
        self.inventory_code = inventory_code
        self.inventory_name = inventory_name
        self.inventory_amount = inventory_amount
        self.inventory_unit = inventory_unit
        self.inventory_location = inventory_location
        self._instances.add(weakref.ref(self))

    @classmethod
    def getinstances(cls):
        dead = set()
        for ref in cls._instances:
            obj = ref()
            if obj is not None:
                yield obj
            else:
                dead.add(ref)
        cls._instances -= dead


for i in range(0, len(inventory_name)):
    globals()[inventory_code2[i]] = InventoryListClass(inventory_code[i], inventory_name[i], inventory_amount[i],
                                                       inventory_unit[i], inventory_location[i])


class SiparisListClass(object):
    _instances = set()

    def __init__(self, OrderCurrent_OrderNo, OrderCurrent_ProductCode, OrderCurrent_ProductName, OrderCurrent_Amount,
                 OrderCurrent_Date):
        self.OrderCurrent_OrderNo = OrderCurrent_OrderNo
        self.OrderCurrent_ProductCode = OrderCurrent_ProductCode
        self.OrderCurrent_ProductName = OrderCurrent_ProductName
        self.OrderCurrent_Amount = OrderCurrent_Amount
        self.OrderCurrent_Date = OrderCurrent_Date
        self._instances.add(weakref.ref(self))

    @classmethod
    def getinstances(cls):
        dead = set()
        for ref in cls._instances:
            obj = ref()
            if obj is not None:
                yield obj
            else:
                dead.add(ref)
        cls._instances -= dead


for i in range(0, len(OrderCurrent_ProductCode)):
    globals()[OrderCurrent_ProductCode2[i]] = SiparisListClass(OrderCurrent_OrderNo[i], OrderCurrent_ProductCode[i],
                                                               OrderCurrent_ProductName[i], OrderCurrent_Amount[i],
                                                               OrderCurrent_Date[i])




