import os, re

ext = "pdf" # change to look for other filetypes

files = os.scandir()
for fn in files:
    if fn.name.endswith(".{}".format(ext)):
        user = re.search("[b|c][0-9]+",fn.name)
        new_fn = "fb_{}.{}".format(user.group(),ext)
        print("{} > {}".format(fn.name,new_fn))
        os.rename(fn.name,new_fn)
