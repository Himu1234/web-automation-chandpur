from bs4 import BeautifulSoup
import urllib3
import lxml

http = urllib3.PoolManager()

url = 'https://newconnection.bpdb.gov.bd/Admin/StatusDetail?status=CENSUS_XEN&zone=101&circle=84&division=326'
response = http.request('GET', url)
soup = BeautifulSoup(response.data, "lxml")
print(soup)
table_body = soup.find('tbody')
rows = table_body.find_all('tr')
print(rows)