my_list = [1, 1, 2, 3, 3, 4, 5, 5, 5, 6, 7, 7]

merged_list = []
index = 0

while index < len(my_list):
    if index == len(my_list) - 1 or my_list[index] != my_list[index + 1]:
        merged_list.append(my_list[index])
    else:
        merged_list.append(my_list[index] + my_list[index + 1])
        index += 1
    index += 1

print(merged_list)
