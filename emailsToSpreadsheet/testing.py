import re

string = "Robinson, Monique <M5ROBINSON@bridgew.edu>"
name = "Alex Sawyer"
string = "asawyer83@hotmail.com"
domain = "hotmail.com"
pattern = '(.*)<(.*)>'

table = {}
table[string] = [name, domain]
table["anfischl@gmail.com"] = ["Mark Fischl", "gmail.com"]
table["mafischl@gmail.com"] = ["Andrew Fischl", "gmail.com"]

for key, value in sorted(table.items(), key=lambda e: e[1], reverse = True):
    print(key, value)
