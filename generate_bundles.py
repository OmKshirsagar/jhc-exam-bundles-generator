import pandas as pd

seating = pd.read_excel("seating_arrangement.xlsx")

# print(seating.iloc[:16, :])

blocks = [block for block in list(seating['ROOMS']) if type(block) == type("A")]
departments = set([department for department in list(seating.iloc[:,1])
               if type(department) == type("A") and not department == 'roll nos'])


room = ""

bundles = list()

addedDepts = list()

for i in range(len(list(seating['ROOMS']))):
    bundle = dict()
    if seating['ROOMS'][i] in blocks:
        room = seating['ROOMS'][i]
    if seating.iloc[i, 1] in departments:
        bundle['department'] = seating.iloc[i, 1]
        bundle['total_no_of_papers'] = seating['SEATS OCCUPIED'][i]
        bundle['block_no'] = room
        addedDepts.append(seating.iloc[i, 1])
        bundle['bundle_no'] = addedDepts.count(seating.iloc[i, 1])
        bundle['total_no_of_bundles'] = addedDepts.count(seating.iloc[i, 1])
        bundle['roll_nos'] = [rolls for rolls in range(int(seating['From'][i+1]), int(seating['To'][i+1]))]
        bundles.append(bundle)

print(blocks)
print(departments)
for bundle in bundles:
    print(bundle)