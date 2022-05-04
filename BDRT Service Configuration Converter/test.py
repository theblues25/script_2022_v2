from netaddr import IPNetwork, IPAddress

nhc = '1.1.1.1'.split()[0]
addr = '1.1.1.0/24'
print(nhc)
if IPAddress(nhc) in IPNetwork(addr):
    print('match')
