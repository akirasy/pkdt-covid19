#!/usr/bin/env python3

import pathlib
from collections import Counter

thefile = pathlib.Path('count.txt')
list_count = thefile.read_text().split('\n')
list_count.sort()
del list_count[0]
test_list = [int(i) for i in list_count]

print(f'Start count value: {test_list[0]}')
print(f'End count value: {test_list[-1]}')

d = Counter(test_list)
#print(d)

new_list = list([item for item in d if d[item]>1])
print(f'Total duplicate found: {len(new_list)}')
for i in new_list:
    print(i)
