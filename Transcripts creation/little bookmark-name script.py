# %%
data_file_path = "./Data/bookmarks name.txt"

bookmark_list = []

with open(data_file_path, 'r') as fh:
    for line in fh:
        bookmark_list.append(line[:-1])  # убираем \n

bookmark_list
# %%
for name in bookmark_list:
    print(name+"_credits")
    print(name+"_mode")
    print(name+"_Academic_results")
    print(name+"_Grades", "\n")
