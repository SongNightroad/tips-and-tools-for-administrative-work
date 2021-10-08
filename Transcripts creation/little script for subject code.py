# Script for making code without changing columns
# %%
data_file_path = "./Data/bookmarks name.txt"

bookmark_list = []

with open(data_file_path, 'r') as fh:
    for line in fh:
        bookmark_list.append(line[:-1])  # убираем \n

bookmark_list
# %%
for i in bookmark_list:
    string1 = "'" + i
    string2 = 'current_column = "T"'
    string3 = 'subject_name = "' + i + '"'
    string4 = "add_subject wd, wdDOC, sh, iRow, filePath_transcripts, exam_mode, current_column, subject_name"
    string5 = "'##########################################"

    print("\t\t" + string1, string2, string3, string4, string5, sep="\n\t\t")
# %%
