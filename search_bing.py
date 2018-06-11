# !/usr/bin/python
# -*- coding: utf-8 -*-

import urllib
import urllib2
import re

url = 'http://cn.bing.com/search'
query = urllib.urlencode({'q':'星空'})
print url + '?' + query
page = urllib.urlopen(url + '?' + query)
content = page.read()
partern = '<li class=.+?><h2><a href="(.+?)".+?>'

i = 0
for match in re.findall(partern, content):
    print match
    i += 1
    urllib.urlretrieve(match, 'd_' + str(i) + '.html')
