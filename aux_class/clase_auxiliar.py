
from collections import namedtuple

Cisco_Articles = namedtuple('Cisco_Articles', 'sku serv_lev')

class CiscoItem:

    def __init__(self, sku, serv_lev, backout='', list_price=0, eos='', smartnet_list=''):
        self.sku = sku
        self.serv_lev = serv_lev
        self.backout = backout
        self.list_price = list_price
        self.eos = eos
        self.smartnet_list = smartnet_list
        self.smartnet_sku = ''
        self.service_price_list = ''
        self.gdc_cost = 0.0

    def __repr__(self):
        return ('sku: {}, sla: {}, back: {}, precio: {}, eos: {}, smartnet: {} {}'.format(self.sku, self.serv_lev,
                                                                                          self.backout, self.list_price,
                                                                                          self.eos,
                                                                                          self.smartnet_sku,
                                                                                          self.smartnet_list))
