#!/usr/bin/env python3

import pathlib

thefile = pathlib.Path('count.txt')
list_count = thefile.read_text().split('\n')
list_count.sort()
del list_count[0]
test_list = [int(i) for i in list_count]

print(f'Start count value: {test_list[0]}')
print(f'End count value: {test_list[-1]}')
missing_elements = []
for ele in range(test_list[0], test_list[-1]+1):
    if ele not in test_list:
        missing_elements.append(ele)
        
print(f'Total number missing: {len(missing_elements)}')
print()
for i in missing_elements:
    print(i)
