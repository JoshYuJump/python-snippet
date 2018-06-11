#!/usr/bin/python

import urllib
import win32com.client 
import re
import time


class db(object):
    '''database class'''
    def __init__(self, db_path):
        self.db_path = db_path
        self.connect_db()

    def connect_db(self):
        self.conn = win32com.client.Dispatch(r'ADODB.Connection')   
        DSN = 'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=' + self.db_path + ';'
        self.conn.Open(DSN)

    def close_db(self):
        self.conn.close()

    def get_rs(self, table_name):
        self.rs = win32com.client.Dispatch(r'ADODB.Recordset')
        self.rs.Open('[' + table_name + ']', self.conn, 1, 3)

    def test_insert_db(self):
        #self.conn.Execute('CREATE TABLE test (name nvarchar(200))')
        self.conn.Execute("INSERT INTO test(name) VALUES('123456')")

    def insert_db(self, o):
        self.rs.AddNew()
        self.rs['Keywords'] = o.keyword
        self.rs['Category1'] = o.category1
        self.rs['Category2'] = o.category2
        self.rs['Category3'] = o.category3
        self.rs['Buyers'] = o.buyers
        self.rs['Clicks'] = o.clicks
        self.rs['Price1'] = o.price1
        self.rs['Price2'] = o.price2
        self.rs.Update()
        print 'Insert Successfully'

    
    def insert_db_category(self, o):
        self.rs.AddNew()
        self.rs['Category'] = o.category
        self.rs.Update()
        print 'Insert Category Successfully'    


class keyword(object):
    def __init__(self):
        pass

def get_keywords(page):
    base_url = 'http://etleida.etbao.cn/nav.php'
    pattern_unit = '<td>([\S\s]*?)<\/td>\s*'
    pattern = re.compile(r'<tr>\s*'+ (pattern_unit * 9) +'</tr>')
    #pattern = re.compile(r'\d+')
    url = ''
    if i == 1:
        url = base_url
    else:
        url = base_url + '?sort=buys-desc&page=' + str(i)
    print url   
    content = urllib.urlopen(url).read().decode('utf-8-sig')
    for m in pattern.finditer(content):
        keyword_instance = keyword()
        keyword_instance.keyword = m.group(2)
        keyword_instance.category1 = m.group(3)
        keyword_instance.category2 = m.group(4)
        keyword_instance.category3 = m.group(5)
        keyword_instance.buyers = m.group(6)
        keyword_instance.clicks = m.group(7)
        keyword_instance.price1 = m.group(8).split('~')[0]
        keyword_instance.price2 = m.group(8).split('~')[1]
        #print keyword_instance.__dict__.items()
        db.insert_db(keyword_instance)

def get_top_category():
    url = 'http://etleida.etbao.cn/nav.php'
    pattern = re.compile(r'<div\sclass="item">\s*<div class="tit">\s*<a.+">(.+)<\/a>')
    content = urllib.urlopen(url).read().decode('utf-8-sig')
    for m in pattern.finditer(content):
        keyword_instance = keyword()
        keyword_instance.category = m.group(1)
        db.insert_db_category(keyword_instance)


if __name__ == '__main__':
    print 'Execute begin'
    db = db('E:/SubwayKeywords.mdb')
    db.get_rs('Categories')


    # collect all keywords
    # for i in range(201, 2001):
    #   get_keywords(i)
    #   print 'Sleep a while.'
    #   time.sleep(0.5)

    # collect top catetory
    get_top_category()
